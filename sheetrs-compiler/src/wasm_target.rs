//! The WebAssembly target: a module that holds a compiled workbook.
//!
//! `sheetc build --target wasm` does not compile a program to WebAssembly code: it completes a
//! *base module* (the runtime, built by the `sheetrs-rt-wasm` crate, which has the engine and no
//! program) with the compiled program, the way a native executable is its runtime plus a payload.
//! A module cannot read its own bytes, so the payload does not go at the end: the base has a slot
//! (a marker followed by an address and a length, both zero), and [`complete`] writes the payload as
//! one more data segment at the end of the memory the base starts with, enlarges that memory to hold
//! it, and writes the address and the length into the slot. The allocator of the module only grows
//! the memory, so nothing it allocates overlaps the payload. The result needs nothing from the
//! host: a browser or Node (the hosts it is tested in) instantiates it and calls its exports.
//!
//! The operation is a splice of three sections (memory, data, and the data count if the base has
//! one) over the bytes of the base; every other section is copied as it is.

use crate::error::{Error, Result};

/// The marker of the slot in the base, and the bytes after it: the address and the length.
const MARKER: &[u8; 16] = b"SHEETRS-WASM-SLT";
const SLOT: usize = 24;
const PAGE: u64 = 65_536;
/// The largest memory a module is given: the payload and what the module allocates while it runs
/// have to fit in a 32-bit address space, with room left for the second.
const LIMIT: u64 = 3 << 30;

fn bad(what: &str) -> Error {
    Error::Wasm(format!(
        "the base module is not one this compiler can complete: {what}"
    ))
}

struct Reader<'a> {
    bytes: &'a [u8],
    at: usize,
}

impl<'a> Reader<'a> {
    fn byte(&mut self) -> Result<u8> {
        let b = *self
            .bytes
            .get(self.at)
            .ok_or_else(|| bad("it ends too soon"))?;
        self.at += 1;
        Ok(b)
    }

    fn leb(&mut self) -> Result<u64> {
        let (mut result, mut shift) = (0u64, 0u32);
        loop {
            let b = self.byte()?;
            if shift >= 64 {
                return Err(bad("a number is too long"));
            }
            result |= u64::from(b & 0x7f) << shift;
            shift += 7;
            if b & 0x80 == 0 {
                return Ok(result);
            }
        }
    }

    fn take(&mut self, n: usize) -> Result<&'a [u8]> {
        let end = self
            .at
            .checked_add(n)
            .ok_or_else(|| bad("a size overflows"))?;
        let slice = self
            .bytes
            .get(self.at..end)
            .ok_or_else(|| bad("it ends too soon"))?;
        self.at = end;
        Ok(slice)
    }
}

fn unsigned(out: &mut Vec<u8>, mut v: u64) {
    loop {
        let b = (v & 0x7f) as u8;
        v >>= 7;
        if v == 0 {
            out.push(b);
            return;
        }
        out.push(b | 0x80);
    }
}

fn signed(out: &mut Vec<u8>, mut v: i64) {
    loop {
        let b = (v & 0x7f) as u8;
        v >>= 7;
        if (v == 0 && b & 0x40 == 0) || (v == -1 && b & 0x40 != 0) {
            out.push(b);
            return;
        }
        out.push(b | 0x80);
    }
}

/// A section: its id and its body.
fn section(out: &mut Vec<u8>, id: u8, body: &[u8]) {
    out.push(id);
    unsigned(out, body.len() as u64);
    out.extend_from_slice(body);
}

/// The base completed with `payload` (see [`sheetrs_engine::bundle::payload`]).
///
/// # Errors
///
/// When `base` is not a module with one memory, active data segments and the slot of the runtime
/// (it is not the base that `sheetrs-rt-wasm` builds), when it already carries a program, or
/// when the payload does not fit the memory a module can have.
pub fn complete(base: &[u8], payload: &[u8]) -> Result<Vec<u8>> {
    if base.get(..8) != Some(b"\0asm\x01\0\0\0") {
        return Err(bad("it is not a WebAssembly module"));
    }
    let mut r = Reader { bytes: base, at: 8 };
    // (id, body) of every section, in order.
    let mut sections: Vec<(u8, &[u8])> = Vec::new();
    while r.at < base.len() {
        let id = r.byte()?;
        let size = usize::try_from(r.leb()?).map_err(|_| bad("a section is too big"))?;
        sections.push((id, r.take(size)?));
    }
    // The memory: one, with no 64-bit addresses. Its size in pages is what the payload goes after.
    let memory_at = sections
        .iter()
        .position(|(id, _)| *id == 5)
        .ok_or_else(|| bad("it defines no memory (is the memory imported?)"))?;
    let mut m = Reader {
        bytes: sections[memory_at].1,
        at: 0,
    };
    if m.leb()? != 1 {
        return Err(bad("it has more than one memory"));
    }
    let flags = m.byte()?;
    if flags & !1 != 0 {
        return Err(bad("its memory is shared or has 64-bit addresses"));
    }
    let min = m.leb()?;
    let max = if flags & 1 == 1 { Some(m.leb()?) } else { None };
    let start = min * PAGE;
    let end = start + payload.len() as u64;
    if end > LIMIT {
        return Err(Error::Wasm(format!(
            "the program is too big for a WebAssembly module: {} bytes of program after {} bytes of memory (the limit is {LIMIT})",
            payload.len(),
            start
        )));
    }
    let pages = end.div_ceil(PAGE);
    let mut memory = Vec::new();
    unsigned(&mut memory, 1);
    if let Some(max) = max {
        memory.push(1);
        unsigned(&mut memory, pages);
        unsigned(&mut memory, max.max(pages));
    } else {
        memory.push(0);
        unsigned(&mut memory, pages);
    }
    // The data: the segments of the base and the payload after them, with the slot written.
    let data_at = sections
        .iter()
        .position(|(id, _)| *id == 11)
        .ok_or_else(|| bad("it has no data section"))?;
    let mut d = Reader {
        bytes: sections[data_at].1,
        at: 0,
    };
    let count = d.leb()?;
    let segments = &sections[data_at].1[d.at..];
    let mut patched = segments.to_vec();
    let mut found = None;
    let mut at = 0;
    while let Some(i) = patched[at..]
        .windows(MARKER.len())
        .position(|w| w == MARKER)
        .map(|i| at + i)
    {
        if found.replace(i).is_some() {
            return Err(bad("the slot of the runtime is there twice"));
        }
        at = i + 1;
    }
    let slot = found.ok_or_else(|| bad("it has no slot for a program"))?;
    let body = patched
        .get_mut(slot..slot + SLOT)
        .ok_or_else(|| bad("the slot is cut short"))?;
    if body[16..].iter().any(|b| *b != 0) {
        return Err(Error::Wasm(
            "this module already carries a program: complete the base, not a completed module"
                .into(),
        ));
    }
    body[16..20].copy_from_slice(&(start as u32).to_le_bytes());
    body[20..24].copy_from_slice(&(payload.len() as u32).to_le_bytes());
    let mut data = Vec::new();
    unsigned(&mut data, count + 1);
    data.extend_from_slice(&patched);
    data.push(0); // an active segment of memory 0
    data.push(0x41); // i32.const
    signed(&mut data, i64::from(start as u32 as i32));
    data.push(0x0b);
    unsigned(&mut data, payload.len() as u64);
    data.extend_from_slice(payload);
    // The module, with the three sections changed.
    let mut out = Vec::with_capacity(base.len() + payload.len() + 64);
    out.extend_from_slice(&base[..8]);
    for (i, (id, body)) in sections.iter().enumerate() {
        match *id {
            5 if i == memory_at => section(&mut out, 5, &memory),
            11 if i == data_at => section(&mut out, 11, &data),
            12 => {
                let mut c = Reader { bytes: body, at: 0 };
                let mut body = Vec::new();
                unsigned(&mut body, c.leb()? + 1);
                section(&mut out, 12, &body);
            }
            _ => section(&mut out, *id, body),
        }
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn leb_roundtrip(v: i64) -> i64 {
        let mut out = Vec::new();
        signed(&mut out, v);
        let (mut result, mut shift, mut last) = (0i64, 0, 0u8);
        for b in out {
            result |= i64::from(b & 0x7f) << shift;
            shift += 7;
            last = b;
        }
        if shift < 64 && last & 0x40 != 0 {
            result |= -1i64 << shift;
        }
        result
    }

    #[test]
    fn signed_numbers_survive_their_encoding() {
        for v in [
            0,
            1,
            -1,
            63,
            64,
            -64,
            -65,
            67_108_864,
            i64::from(i32::MIN),
            i64::from(i32::MAX),
        ] {
            assert_eq!(leb_roundtrip(v), v, "{v}");
        }
    }

    #[test]
    fn what_is_not_a_base_is_refused() {
        assert!(complete(b"nope", b"x").is_err());
        assert!(complete(b"\0asm\x01\0\0\0", b"x").is_err());
    }
}

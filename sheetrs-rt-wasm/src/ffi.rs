//! The exports of the module: a C-like interface over linear memory that any host can call.
//!
//! The host writes a request into memory it gets from [`sheetrs_alloc`], calls a function and reads
//! the answer from the *result* buffer ([`sheetrs_result_ptr`], [`sheetrs_result_len`]) or, for a
//! workbook, the *workbook* buffer. A function returns `0` when it worked and the buffer holds
//! the answer, or `1` when it failed and the result buffer holds the message (UTF-8). The buffers
//! are valid until the next call. A host that wants the clock of the world sets it with
//! [`sheetrs_set_clock`] (the module has no clock of its own: without it `NOW()` is the Unix epoch).
//!
//! | Export | |
//! |---|---|
//! | `memory` | the linear memory |
//! | `sheetrs_abi() -> i32` | the version of this interface |
//! | `sheetrs_alloc(len) -> ptr`, `sheetrs_free(ptr, len)` | memory for a request |
//! | `sheetrs_set_clock(seconds: f64)`, `sheetrs_set_seed(seed: f64)` | the clock and the random seed (an integer up to 2^53) |
//! | `sheetrs_meta() -> status` | the `--meta` of an executable, in the result buffer |
//! | `sheetrs_info() -> status` | the `--info` |
//! | `sheetrs_calculate(ptr, len) -> status` | the JSON answer to a request |
//! | `sheetrs_write(ptr, len, with_json) -> status` | the workbook with the answer written, in the workbook buffer, and the JSON in the result buffer if `with_json` (modules built `--embed-template` or `--recalculate`, from the base with the writer) |
//! | `sheetrs_result_ptr() -> ptr`, `sheetrs_result_len() -> len` | the result buffer |
//! | `sheetrs_workbook_ptr() -> ptr`, `sheetrs_workbook_len() -> len` | the workbook buffer |
//! | `sheetrs_has_writer() -> i32` | `1` when the module can write a workbook |

use super::{ABI_VERSION, SHEETRS_SLOT, SLOT_MARKER, book};
use sheetrs_engine::calculate::Overrides;
use std::cell::{Cell, RefCell};

thread_local! {
    static RESULT: RefCell<Vec<u8>> = const { RefCell::new(Vec::new()) };
    static WORKBOOK: RefCell<Vec<u8>> = const { RefCell::new(Vec::new()) };
    static CLOCK: Cell<f64> = const { Cell::new(0.0) };
    static SEED: Cell<Option<u64>> = const { Cell::new(None) };
}

fn clock() -> f64 {
    CLOCK.with(Cell::get)
}

fn overrides() -> Overrides {
    sheetrs_engine::clock::set_wall_clock(clock);
    Overrides {
        now: None,
        rand_seed: SEED.with(Cell::get),
        ..Overrides::default()
    }
}

/// The payload `sheetc` put in the module, from the slot.
pub(crate) fn payload() -> Result<&'static [u8], String> {
    let mut slot = [0u8; 24];
    for (i, byte) in slot.iter_mut().enumerate() {
        // SAFETY: `i` is inside the static; the read is volatile so that the value `sheetc` wrote
        // into the data of the module is read, not the one the compiler saw.
        *byte = unsafe {
            std::ptr::read_volatile(std::ptr::addr_of!(SHEETRS_SLOT).cast::<u8>().add(i))
        };
    }
    if &slot[..16] != SLOT_MARKER {
        return Err("the slot of the module is damaged".into());
    }
    let word = |at: usize| u32::from_le_bytes([slot[at], slot[at + 1], slot[at + 2], slot[at + 3]]);
    let (ptr, len) = (word(16) as usize, word(20) as usize);
    if len == 0 {
        return Err(
            "no program in this module: it is the base that `sheetc build --target wasm` completes"
                .into(),
        );
    }
    // SAFETY: `sheetc` wrote the address and the length of a data segment of this module, which is
    // never freed or written.
    Ok(unsafe { std::slice::from_raw_parts(ptr as *const u8, len) })
}

fn answer(result: Result<String, String>) -> i32 {
    let (status, text) = match result {
        Ok(text) => (0, text),
        Err(message) => (1, message),
    };
    RESULT.with(|r| *r.borrow_mut() = text.into_bytes());
    status
}

fn request<'a>(ptr: *const u8, len: usize) -> &'a [u8] {
    if len == 0 {
        return &[];
    }
    // SAFETY: the host passes a region it got from `sheetrs_alloc` and wrote.
    unsafe { std::slice::from_raw_parts(ptr, len) }
}

#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_abi() -> i32 {
    ABI_VERSION as i32
}

#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_has_writer() -> i32 {
    i32::from(cfg!(feature = "template"))
}

#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_alloc(len: usize) -> *mut u8 {
    let mut v = Vec::<u8>::with_capacity(len.max(1));
    let ptr = v.as_mut_ptr();
    std::mem::forget(v);
    ptr
}

/// # Safety
///
/// `ptr` and `len` are what `sheetrs_alloc(len)` gave, freed once.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn sheetrs_free(ptr: *mut u8, len: usize) {
    // SAFETY: the caller gives back what `sheetrs_alloc` returned, with its capacity.
    drop(unsafe { Vec::from_raw_parts(ptr, 0, len.max(1)) });
}

#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_set_clock(seconds: f64) {
    CLOCK.with(|c| c.set(seconds));
}

#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_set_seed(seed: f64) {
    SEED.with(|s| s.set((seed.is_finite() && seed >= 0.0).then_some(seed as u64)));
}

#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_meta() -> i32 {
    answer(book().map(super::Book::meta))
}

#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_info() -> i32 {
    answer(book().map(super::Book::info))
}

#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_calculate(ptr: *const u8, len: usize) -> i32 {
    answer(book().and_then(|b| b.calculate(request(ptr, len), &overrides())))
}

#[cfg(feature = "template")]
#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_write(ptr: *const u8, len: usize, with_json: i32) -> i32 {
    let done = book().and_then(|b| b.write(request(ptr, len), with_json != 0, &overrides()));
    match done {
        Ok((workbook, json)) => {
            WORKBOOK.with(|w| *w.borrow_mut() = workbook);
            answer(Ok(json.unwrap_or_default()))
        }
        Err(message) => answer(Err(message)),
    }
}

#[cfg(not(feature = "template"))]
#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_write(_ptr: *const u8, _len: usize, _with_json: i32) -> i32 {
    answer(Err(
        "this module cannot write a workbook: it was made from the base without the writer".into(),
    ))
}

#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_result_ptr() -> *const u8 {
    RESULT.with(|r| r.borrow().as_ptr())
}

#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_result_len() -> usize {
    RESULT.with(|r| r.borrow().len())
}

#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_workbook_ptr() -> *const u8 {
    WORKBOOK.with(|w| w.borrow().as_ptr())
}

#[unsafe(no_mangle)]
pub extern "C" fn sheetrs_workbook_len() -> usize {
    WORKBOOK.with(|w| w.borrow().len())
}

//! The runtime of a compiled workbook as a WebAssembly module.
//!
//! `sheetc build --target wasm` takes the module this crate builds (the *base*), adds the compiled
//! program to it and writes the result: a module that holds one workbook and answers requests, with
//! no JavaScript runtime, no `wasm-bindgen` and no imports: a browser or Node (the hosts it is
//! tested in) can instantiate it and call its exports (see [`ffi`]).
//!
//! # How the module finds its program
//!
//! A module cannot read its own bytes, so the program cannot ride at the end like the payload of a
//! native executable. The base has a *slot*, [`SHEETRS_SLOT`]: a marker followed by the address and
//! the length of the payload, both zero. `sheetc` finds the marker in the data of the base, writes
//! the address and the length into it and adds the payload as one more data segment, at the end of
//! the memory the base starts with (which it enlarges to hold it). The allocator of the module only
//! ever grows the memory, so nothing it allocates can overlap the payload. At the first call the
//! module reads the slot and decodes the program from the payload in place.
//!
//! The payload is the one a native executable carries (see
//! [`sheetrs_engine::bundle::payload`]): the program, the blocks of its constants, the template
//! workbook if there is one, and the trailer.

use sheetrs_engine::bundle::{self, Embedded};
use sheetrs_engine::calculate::Overrides;
use sheetrs_engine::engine::ConstSource;
use std::cell::RefCell;

pub mod ffi;

/// What the base module holds before `sheetc` completes it: [`SLOT_MARKER`], then the address and the
/// length of the payload as two little-endian 32-bit integers.
#[unsafe(no_mangle)]
#[used]
pub static SHEETRS_SLOT: [u8; 24] = *b"SHEETRS-WASM-SLT\0\0\0\0\0\0\0\0";

/// The bytes that mark the slot; also what `sheetc` looks for in the base.
pub const SLOT_MARKER: &[u8; 16] = b"SHEETRS-WASM-SLT";

/// The version of the interface of the module (the exports of [`ffi`]).
pub const ABI_VERSION: u32 = 1;

/// A workbook compiled into the module.
pub struct Book {
    embedded: Embedded,
}

impl Book {
    /// Decode the program of `payload`.
    ///
    /// # Errors
    ///
    /// When the payload is not one `sheetc build --target wasm` writes.
    pub fn from_payload(payload: &'static [u8]) -> Result<Self, String> {
        bundle::read_payload(payload)
            .map(|embedded| Self { embedded })
            .ok_or_else(|| "the program of this module is damaged or is not there".into())
    }

    fn base(&self) -> Option<&dyn ConstSource> {
        self.embedded.consts.as_ref().map(|c| c as &dyn ConstSource)
    }

    /// The `--meta` of any executable: the inputs, the outputs and the metadata of the program.
    #[must_use]
    pub fn meta(&self) -> String {
        sheetrs_engine::cli::meta_json(&self.embedded.program).to_string()
    }

    /// The `--info` of any executable.
    #[must_use]
    pub fn info(&self) -> String {
        sheetrs_engine::cli::describe(&self.embedded.program, self.embedded.consts.as_ref())
    }

    /// The JSON answer (`{"valueRanges": [...]}`) to the request in `request` (`{"inputs": {...}, "ranges": [...]}`; empty
    /// for the declared outputs).
    ///
    /// # Errors
    ///
    /// The message of the failure: a request that is not valid for the program, or a run that fails.
    pub fn calculate(&self, request: &[u8], overrides: &Overrides) -> Result<String, String> {
        let request = parse_request(request)?;
        let answer = self
            .embedded
            .program
            .handle_over(self.base(), &request, overrides)
            .map_err(|e| e.to_string())?;
        Ok(answer.to_json().to_string())
    }

    /// The workbook the program was made from with the answer written into it, and the JSON
    /// answer as well when `with_json`.
    ///
    /// # Errors
    ///
    /// When the module carries no workbook (it was built without `--embed-template`), the request
    /// is not valid or the workbook cannot be written.
    #[cfg(feature = "template")]
    pub fn write(
        &self,
        request: &[u8],
        with_json: bool,
        overrides: &Overrides,
    ) -> Result<(Vec<u8>, Option<String>), String> {
        let Some(template) = &self.embedded.template else {
            return Err(
                "this module carries no workbook: build it with --embed-template to write one"
                    .into(),
            );
        };
        let source = sheetrs_io::Book::from_bytes(template.bytes.clone())
            .map_err(|e| format!("the template: {e}"))?;
        let program = &self.embedded.program;
        let request = parse_request(request)?;
        let mut workbook = Vec::new();
        if with_json {
            let (_, answer) = sheetrs_answer::fill_template_to_with_answer(
                &source,
                program,
                self.base(),
                &request,
                overrides,
                &mut workbook,
            )
            .map_err(|e| format!("{e:#}"))?;
            Ok((workbook, Some(answer.to_json().to_string())))
        } else {
            sheetrs_answer::fill_template_to(
                &source,
                program,
                self.base(),
                &request,
                overrides,
                &mut workbook,
            )
            .map_err(|e| format!("{e:#}"))?;
            Ok((workbook, None))
        }
    }

    /// Whether the module carries a workbook to write answers into.
    #[must_use]
    pub fn has_template(&self) -> bool {
        self.embedded.template.is_some()
    }
}

fn parse_request(request: &[u8]) -> Result<serde_json::Value, String> {
    if request.iter().all(u8::is_ascii_whitespace) {
        return Ok(serde_json::Value::Object(serde_json::Map::new()));
    }
    serde_json::from_slice(request).map_err(|e| format!("the request is not JSON: {e}"))
}

thread_local! {
    static BOOK: RefCell<Option<Result<&'static Book, String>>> = const { RefCell::new(None) };
}

/// The book of this module, decoded the first time it is asked for. The payload is read from the slot.
pub(crate) fn book() -> Result<&'static Book, String> {
    BOOK.with(|cell| {
        cell.borrow_mut()
            .get_or_insert_with(|| {
                let payload = ffi::payload()?;
                // The book lives as long as the module: it is decoded once.
                Book::from_payload(payload).map(|b| &*Box::leak(Box::new(b)))
            })
            .clone()
    })
}

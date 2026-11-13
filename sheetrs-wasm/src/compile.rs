//! `sheetc build`: the WebAssembly module of a workbook.
//!
//! In this environment the only target is WebAssembly: the module is the base runtime (the one
//! `sheetrs-rt-wasm` builds, which holds the engine) with the compiled program added to it, and a
//! browser or Node (the hosts it is tested in; it imports nothing) instantiates it. The package carries the two base modules inside
//! (the feature `bundled-runtime`, which `scripts/package-wasm.sh` sets): there is one package,
//! and nothing to fetch to compile.

use crate::book::Book;
use crate::calculate::call_of;
use crate::js::{bytes, compiler_diagnostics, from_js, object, to_js};
use serde_json::Value as Json;
use sheetrs_compiler::memory::BookCompile;
use sheetrs_compiler::memory::Runtimes;
use wasm_bindgen::prelude::*;

static BUNDLED_VALUES: &[u8] = include_bytes!(concat!(env!("SHEETRS_RT_DIR"), "/sheetrs-rt.wasm"));
static BUNDLED_TEMPLATE: &[u8] =
    include_bytes!(concat!(env!("SHEETRS_RT_DIR"), "/sheetrs-rt-template.wasm"));

#[wasm_bindgen]
impl Book {
    /// The WebAssembly module of the workbook: what `sheetc build FILE --target wasm` writes.
    /// `options` are the keys of a `sheetc.toml`, `declarations`, and `embed_template`, except
    /// `json`, which is chosen when the module runs; the target is not an option: a request for
    /// another one is an error.
    ///
    /// Returns `{module, cells, reachable, formulas, folded, kept, cycles, diagnostics}`; `module` is
    /// a `Uint8Array` whose exports are described in the README of the package. `diagnostics` are the
    /// warnings of the build as records (`{level, code, message, context}`), also handed to
    /// `onDiagnostic`.
    ///
    /// # Errors
    ///
    /// A message when operations are pending, the workbook cannot be read or compiled, the options
    /// are not valid, or the target is not `wasm`.
    pub fn compile(&self, options: JsValue) -> Result<JsValue, JsValue> {
        self.queryable()?;
        let mut value: Json = if options.is_undefined() || options.is_null() {
            Json::Null
        } else {
            from_js(options, "The options")?
        };
        if let Some(target) = value.as_object_mut().and_then(|o| o.remove("target")) {
            return Err(JsValue::from_str(&format!(
                "Compile error: there is one target here, WebAssembly: `target` ({target}) is not an option of this environment"
            )));
        }
        if value
            .as_object()
            .is_some_and(|o| o.get("json").is_some_and(|j| !j.is_null()))
        {
            return Err(JsValue::from_str(
                "Compile error: `json` is chosen when the module runs (`sheetrs_write` with or without JSON), not when it is built",
            ));
        }
        let call = call_of(to_js(&value)?)?;
        let mode = call.mode;
        // Only the base this mode needs is used.
        let (values, template): (&[u8], &[u8]) = if mode.is_workbook() {
            (&[], BUNDLED_TEMPLATE)
        } else {
            (BUNDLED_VALUES, &[])
        };
        let (module, report) = self
            .inner
            .build_wasm(&call.options, mode, Runtimes { values, template })
            .map_err(|e| JsValue::from_str(&format!("Compile error: {e}")))?;
        let diagnostics = compiler_diagnostics(&report.warnings);
        self.emit(&diagnostics)?;
        object(&[
            ("module", bytes(&module)),
            ("cells", JsValue::from_f64(report.cells_total as f64)),
            (
                "reachable",
                JsValue::from_f64(report.cells_reachable as f64),
            ),
            ("formulas", JsValue::from_f64(report.formulas_total as f64)),
            ("folded", JsValue::from_f64(report.formulas_folded as f64)),
            ("kept", JsValue::from_f64(report.formulas_kept as f64)),
            ("cycles", JsValue::from_f64(report.cycles as f64)),
            ("diagnostics", to_js(&diagnostics)?),
        ])
    }
}

//! WebAssembly bindings of the `SheetRS` suite, for a browser or Node (`--target web`).
//!
//! The package has one object, `Book`: a workbook made from its bytes, whose format is
//! the one of the content. It reads nothing when it is made; `format` comes from the content, and `stats`, `lint`, `query`,
//! `range`, `find` and `inspect` read what they need; the operations (`deleteSheets`, `deleteNames`,
//! `formulaToValue`, `deleteRows`, `deleteColumns`, `replace`, `calculate`) accumulate until `save`
//! runs them in the order the planner chooses; `compile` makes a WebAssembly module of the
//! workbook (with the base modules inside the package). The features `query`, `edit` and `calculate`
//! choose what the module holds. The methods, their options and the exports of a compiled module
//! are in `sheetrs-wasm/README.md`.

mod book;
#[cfg(feature = "calculate")]
mod calculate;
#[cfg(feature = "bundled-runtime")]
mod compile;
#[cfg(feature = "edit")]
mod edit;
mod js;
#[cfg(feature = "query")]
mod query;

use wasm_bindgen::prelude::*;

/// The wall clock of the engine is `Date.now()`: `std::time::SystemTime::now()` panics on
/// WebAssembly. It runs when the module is instantiated, before any call.
#[cfg(feature = "calculate")]
#[wasm_bindgen(start)]
fn start() {
    fn now() -> f64 {
        js_sys::Date::now() / 1000.0
    }
    sheetrs_engine::clock::set_wall_clock(now);
}

/// Makes a panic of the module show its message in the console of the page.
#[wasm_bindgen(js_name = initPanicHook)]
pub fn init_panic_hook() {
    console_error_panic_hook::set_once();
}

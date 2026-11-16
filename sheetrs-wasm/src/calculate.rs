//! `sheetc`: what the compiler says about a workbook (`inspect`) and one formula (`evaluateFormula`), in
//! memory, with the engine that reproduces `LibreOffice`.

use crate::book::Book;
use crate::js::{from_js, to_js};
use serde_json::Value as Json;
use sheetrs_compiler::memory::{self, BookCompile, Options};
use sheetrs_engine::ir::Mode;

use wasm_bindgen::prelude::*;

/// The options of a call: the keys of a `sheetc.toml` (`strict`, `now`, `seed`, `host`,
/// `enable_features`, `locale`, `settings`), `declarations` (the JSON of a `settings_file`: the
/// inputs and outputs of a workbook with no settings sheet) and, as the flags of the command do,
/// `embed_template` (the workbook is the answer, as `-o OUT` of `sheetc calculate`) and `json` (which
/// only asks for the JSON answer besides it, and is not valid without it).
#[cfg_attr(not(feature = "bundled-runtime"), allow(dead_code))]
pub(crate) struct Call {
    pub(crate) options: Options,
    pub(crate) mode: Mode,
}

pub(crate) fn call_of(options: JsValue) -> Result<Call, JsValue> {
    let mut value: Json = if options.is_undefined() || options.is_null() {
        Json::Null
    } else {
        from_js(options, "The options")?
    };
    let mut flag = |name: &str| -> Result<bool, JsValue> {
        match value.as_object_mut().and_then(|o| o.remove(name)) {
            None | Some(Json::Null) => Ok(false),
            Some(Json::Bool(b)) => Ok(b),
            Some(_) => Err(JsValue::from_str(&format!("`{name}` is true or false"))),
        }
    };
    let embed = flag("embed_template")?;
    let json = flag("json")?;
    let mode = if embed { Mode::Embedded } else { Mode::Values };
    if !mode.is_workbook() && json {
        return Err(JsValue::from_str(
            "`json` needs `embed_template`: without it the answer is JSON already",
        ));
    }
    let options = Options::from_json(&value).map_err(|e| JsValue::from_str(&e.to_string()))?;
    Ok(Call { options, mode })
}

fn failure(e: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&format!("Calculate error: {e}"))
}

#[wasm_bindgen]
impl Book {
    /// What `sheetc inspect` says: `{sheets, date_system, cells, formulas, circular_groups,
    /// run_time_references, unsupported_functions, functions: [{name, calls, implemented}],
    /// diagnostics}`. `options` are the keys of a `sheetc.toml` (only the configuration is used).
    ///
    /// # Errors
    ///
    /// A message when operations are pending, the workbook cannot be read or compiled, or the
    /// options are not valid.
    pub fn inspect(&self, options: JsValue) -> Result<Inspection, JsValue> {
        self.queryable()?;
        let call = call_of(options)?;
        let report = self.inner.inspect(&call.options).map_err(failure)?;
        Ok(Inspection { inner: report })
    }
}

/// What the compiler says about a workbook.
#[wasm_bindgen]
pub struct Inspection {
    inner: memory::Inspection,
}

#[wasm_bindgen]
impl Inspection {
    /// The report as an object.
    ///
    /// # Errors
    ///
    /// A message when the result cannot be built.
    #[wasm_bindgen(js_name = toJSON)]
    pub fn to_json(&self) -> Result<JsValue, JsValue> {
        to_js(&self.inner)
    }
}

/// `sheetc evaluate`: the value of one formula, with the regional settings of `locale` (a BCP-47
/// tag, `en` by default). `options` may carry `now` and `seed`.
///
/// # Errors
///
/// A message when the locale or the formula is not valid.
// `wasm_bindgen` takes an optional text by value.
#[allow(clippy::needless_pass_by_value)]
#[wasm_bindgen(js_name = evaluateFormula)]
pub fn evaluate_formula(
    formula: &str,
    locale: Option<String>,
    options: JsValue,
) -> Result<JsValue, JsValue> {
    let call = call_of(options)?;
    let value = memory::evaluate(formula, locale.as_deref(), &call.options).map_err(failure)?;
    to_js(&value)
}

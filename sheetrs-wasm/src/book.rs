//! `Book`: the workbook, the one object of the package.
//!
//! A `Book` is made from the bytes of a workbook (`new Book(bytes)`), whose format is the one of the
//! content. It reads nothing when it is made and keeps nothing of the workbook: what it holds is its
//! bytes and the operations to make to it. Everything is read when it is asked for, from the bytes
//! the `Book` keeps: `range`, `find`, `lint` and `stats` read what they need, and the operations
//! (`deleteSheets`, `deleteNames`, `formulaToValue`, `deleteRows`, `deleteColumns`, `replace`,
//! `calculate`) only accumulate in the plan until `save` runs them. The operations are not a sequence:
//! each is written against the workbook as it is, and the planner decides the order.

// `wasm_bindgen` takes an optional text by value.
#![allow(clippy::needless_pass_by_value)]

use crate::js::to_js;
use serde::Serialize;
use sheetrs_io::Book as Inner;
use sheetrs_linter::{BookLint, Linter, LinterConfig};
use sheetrs_stats::{BookStats, WorkbookStats};
use std::rc::Rc;
use wasm_bindgen::prelude::*;

/// A workbook: its bytes, and the operations to make to it.
#[wasm_bindgen]
pub struct Book {
    pub(crate) inner: Inner,
    pub(crate) on_diagnostic: Option<js_sys::Function>,
    pub(crate) dates: bool,
    #[cfg(feature = "edit")]
    pub(crate) report: Option<crate::edit::Report>,
}

pub(crate) fn failure(prefix: &str, e: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&format!("{prefix}: {e}"))
}

#[wasm_bindgen]
impl Book {
    /// A workbook from its bytes (a `Uint8Array`). The format (XLSX, XLSM or ODS) is the one of the
    /// content. It is read once to count what it holds; see the properties.
    ///
    /// # Errors
    ///
    /// A message when the bytes are not an XLSX, XLSM or ODS workbook.
    #[wasm_bindgen(constructor)]
    pub fn new(data: &[u8]) -> Result<Book, JsValue> {
        let inner = Inner::from_bytes(data).map_err(|e| failure("Reader error", e))?;
        Ok(Book {
            inner,
            on_diagnostic: None,
            dates: false,
            #[cfg(feature = "edit")]
            report: None,
        })
    }

    /// `"xlsx"`, `"xlsm"` or `"ods"`.
    #[wasm_bindgen(getter)]
    pub fn format(&self) -> String {
        self.inner.format().extension().to_string()
    }

    /// The statistics of the workbook, as `sheetrs stats` reports them. It reads the workbook, in a
    /// single streaming pass, each time it is called; the `Book` keeps nothing of what it found.
    ///
    /// # Errors
    ///
    /// A message when the workbook cannot be read.
    pub fn stats(&self) -> Result<Stats, JsValue> {
        let stats = self.inner.stats().map_err(|e| failure("Reader error", e))?;
        Ok(Stats {
            inner: Rc::new(stats),
        })
    }

    /// Whether `range`, `find`, `lint` and `query` give dates and times as typed values
    /// (`{type: "date" | "time" | "datetime", value}`) and not as serial numbers (the default).
    #[wasm_bindgen(getter)]
    pub fn dates(&self) -> bool {
        self.dates
    }

    #[wasm_bindgen(setter)]
    pub fn set_dates(&mut self, dates: bool) {
        self.dates = dates;
    }

    /// A function that receives each diagnostic of a run (`{level, code, message, context}`) after it
    /// ends; an exception it throws ends the call. `null` removes it.
    #[wasm_bindgen(setter = onDiagnostic)]
    pub fn set_on_diagnostic(&mut self, callback: Option<js_sys::Function>) {
        self.on_diagnostic = callback;
    }

    /// Lints the workbook with the rules of `config_toml` (the text of a `sheetrs.toml`; the
    /// default rules when empty or absent).
    ///
    /// # Errors
    ///
    /// A message when operations are pending, the workbook cannot be read or the configuration is
    /// not valid.
    pub fn lint(&self, config_toml: Option<String>) -> Result<LintReport, JsValue> {
        self.queryable()?;
        let config = match config_toml.as_deref().map(str::trim) {
            None | Some("") => LinterConfig::default(),
            Some(text) => LinterConfig::from_toml(text).map_err(|e| failure("Config error", e))?,
        };
        let report = self
            .inner
            .lint(&Linter::with_config(config))
            .map_err(|e| failure("Reader error", e))?;
        Ok(LintReport { inner: report })
    }

    /// A query is about the workbook as the bytes have it: it cannot run while operations are
    /// pending.
    pub(crate) fn queryable(&self) -> Result<(), JsValue> {
        #[cfg(feature = "edit")]
        if !self.inner.operations().is_empty() {
            return Err(JsValue::from_str(&format!(
                "Pending operations: {} operation(s) have not run, so a query would be about a workbook that does not exist yet; call save() first",
                self.inner.operations().len()
            )));
        }
        Ok(())
    }
}

impl Book {
    /// Hands each diagnostic to `onDiagnostic`, if there is one.
    #[cfg(feature = "edit")]
    pub(crate) fn emit(&self, list: &[sheetrs_protocol::Diagnostic]) -> Result<(), JsValue> {
        if let Some(callback) = &self.on_diagnostic {
            for diagnostic in list {
                callback.call1(&JsValue::NULL, &to_js(diagnostic)?)?;
            }
        }
        Ok(())
    }
}

/// The statistics of a workbook.
#[wasm_bindgen]
pub struct Stats {
    inner: Rc<WorkbookStats>,
}

#[wasm_bindgen]
impl Stats {
    /// Whether the workbook carries macros (found, never run).
    #[wasm_bindgen(getter, js_name = hasMacros)]
    pub fn has_macros(&self) -> bool {
        self.inner.has_macros
    }

    /// The defined names and what each refers to, as the file writes it (`{"Rate": "Data!$B$2"}`).
    ///
    /// # Errors
    ///
    /// A message when the result cannot be built.
    #[wasm_bindgen(getter, js_name = definedNames)]
    pub fn defined_names(&self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.defined_names)
    }

    /// The totals and the sheets, as `sheetrs stats --format json` writes them (`total_file_size` is
    /// 0: bytes carry no size).
    ///
    /// # Errors
    ///
    /// A message when the result cannot be built.
    #[wasm_bindgen(js_name = toJSON)]
    pub fn to_json(&self) -> Result<JsValue, JsValue> {
        to_js(&*self.inner)
    }
}

/// The result of linting a workbook.
#[wasm_bindgen]
pub struct LintReport {
    inner: sheetrs_linter::LintReport,
}

#[wasm_bindgen]
impl LintReport {
    /// The number of violations.
    #[wasm_bindgen(getter)]
    pub fn count(&self) -> usize {
        self.inner.violations().len()
    }

    /// The violations with everything resolved: `rule_id`, `rule_name`, `severity`, `message`,
    /// `scope`, `sheet`, `cell`, `row`, `col`.
    ///
    /// # Errors
    ///
    /// A message when the result cannot be built.
    #[wasm_bindgen(getter)]
    pub fn records(&self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.records())
    }

    /// The report for a person, as text.
    #[wasm_bindgen(js_name = toText)]
    pub fn to_text(&self, file: Option<String>) -> String {
        self.inner.to_text(file.as_deref())
    }

    /// The JSON document of `sheetrs lint --format json` (`file`, `summary`, `violations`) as an
    /// object, so that `JSON.stringify(report)` gives it.
    ///
    /// # Errors
    ///
    /// A message when the document cannot be built.
    #[wasm_bindgen(js_name = toJSON)]
    pub fn to_json(&self) -> Result<JsValue, JsValue> {
        let value: serde_json::Value = serde_json::from_str(&self.inner.to_json(None))
            .map_err(|e| failure("Serialization error", e))?;
        to_js(&value)
    }

    /// The same document as text, exactly as `sheetrs lint --format json` writes it; `file` is the
    /// name of the workbook it refers to.
    #[wasm_bindgen(js_name = toJsonText)]
    pub fn to_json_text(&self, file: Option<String>) -> String {
        self.inner.to_json(file.as_deref())
    }

    /// The SARIF 2.1.0 log as text, as `sheetrs lint --format sarif` writes it; `file` is the name of
    /// the workbook it refers to (its `artifactLocation`).
    #[wasm_bindgen(js_name = toSarif)]
    pub fn to_sarif(&self, file: Option<String>) -> String {
        self.inner.to_sarif(file.as_deref())
    }
}

#[derive(Serialize)]
struct RuleInfo {
    id: String,
    name: String,
    category: String,
    is_default: bool,
}

/// The rules of the linter: `{id, name, category, is_default}`.
///
/// # Errors
///
/// A message when the result cannot be serialized.
#[wasm_bindgen(js_name = lintRules)]
pub fn lint_rules() -> Result<JsValue, JsValue> {
    let config = LinterConfig::default();
    let rules: Vec<RuleInfo> = sheetrs_linter::rules::registry::rule_metadata(&config)
        .iter()
        .map(|m| RuleInfo {
            id: m.id.to_string(),
            name: m.name.clone(),
            category: m.category.as_str().to_string(),
            is_default: m.is_default,
        })
        .collect();
    to_js(&rules)
}

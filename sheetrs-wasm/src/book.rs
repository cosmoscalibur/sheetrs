//! `Book`: the workbook, the one object of the package.
//!
//! A `Book` is made from the bytes of a workbook (`new Book(bytes)`), whose format is the one of the
//! content. It reads nothing when it is made: the first shortcut property that describes the workbook
//! reads it once, in a single streaming pass that counts what it holds (sheets, used cells, formulas,
//! names, macros), and keeps only those numbers: they are the shortcut properties (`format`, `sheetNames`, `sheets`, `definedNames`, `dateSystem`,
//! `hasMacros`, the eleven `total*`). Everything else is read when it is asked for, from the bytes
//! the `Book` keeps: `range`, `find` and `lint` read what they need, and the operations
//! (`deleteSheets`, `deleteNames`, `formulaToValue`, `deleteRows`, `deleteColumns`, `replace`,
//! `calculate`) only accumulate in the plan until `save` runs them. The operations are not a sequence:
//! each is written against the workbook as it is, and the planner decides the order.

// `wasm_bindgen` takes an optional text by value.
#![allow(clippy::needless_pass_by_value)]

use crate::js::to_js;
use serde::Serialize;
use sheetrs_io::Book as Inner;
use sheetrs_io::reader::SheetState;
use sheetrs_linter::{BookLint, Linter, LinterConfig};
use sheetrs_stats::{BookStats, WorkbookStats};
use std::cell::OnceCell;
use std::rc::Rc;
use wasm_bindgen::prelude::*;

/// What the first read that describes the workbook found out.
struct Opened {
    stats: Rc<WorkbookStats>,
    /// The null date as `YYYY-MM-DD`, or the reason the workbook has none the package supports.
    date_system: Result<String, String>,
}

/// A workbook: its bytes, and the operations to make to it.
#[wasm_bindgen]
pub struct Book {
    pub(crate) inner: Inner,
    /// Made the first time a property describes the workbook, with the reason if it cannot be read.
    opened: OnceCell<Result<Opened, String>>,
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
            opened: OnceCell::new(),
            on_diagnostic: None,
            dates: false,
            #[cfg(feature = "edit")]
            report: None,
        })
    }

    /// What the first read that describes the workbook found out, made now if it was not.
    fn described(&self) -> Result<&Opened, JsValue> {
        self.opened
            .get_or_init(|| {
                let stats = self
                    .inner
                    .stats()
                    .map_err(|e| format!("Reader error: {e}"))?;
                let date_system = self
                    .inner
                    .date_system()
                    .map(|system| {
                        let (year, month, day) = system.null_date();
                        format!("{year:04}-{month:02}-{day:02}")
                    })
                    .map_err(|e| format!("{e:#}"));
                Ok(Opened {
                    stats: Rc::new(stats),
                    date_system,
                })
            })
            .as_ref()
            .map_err(|e| JsValue::from_str(e))
    }

    /// `"xlsx"`, `"xlsm"` or `"ods"`.
    #[wasm_bindgen(getter)]
    pub fn format(&self) -> String {
        self.inner.format().extension().to_string()
    }

    /// The date of serial 0 as `YYYY-MM-DD`: `1899-12-30`, `1900-01-01` or `1904-01-01`.
    ///
    /// # Errors
    ///
    /// A message when the workbook has another date system.
    #[wasm_bindgen(getter, js_name = dateSystem)]
    pub fn date_system(&self) -> Result<String, JsValue> {
        self.described()?
            .date_system
            .clone()
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Whether the workbook carries macros (found, never run).
    #[wasm_bindgen(getter, js_name = hasMacros)]
    pub fn has_macros(&self) -> Result<bool, JsValue> {
        Ok(self.described()?.stats.has_macros)
    }

    /// The names of the sheets, in order.
    #[wasm_bindgen(getter, js_name = sheetNames)]
    pub fn sheet_names(&self) -> Result<Vec<String>, JsValue> {
        Ok(self
            .described()?
            .stats
            .sheets
            .iter()
            .map(|s| s.name.clone())
            .collect())
    }

    /// The sheets by name (`book.sheets["Data"]`): `name`, `usedRange` (`"A1:C9"` or `null`),
    /// `state` (`"visible"`, `"hidden"` or `"very_hidden"`) and the counts `usedCells`, `formulas`,
    /// `errors`, `emptyCells`, `comments`, `charts`, `tables` and `xmlSize`.
    ///
    /// # Errors
    ///
    /// A message when the result cannot be built.
    #[wasm_bindgen(getter)]
    pub fn sheets(&self) -> Result<JsValue, JsValue> {
        #[derive(Serialize)]
        #[serde(rename_all = "camelCase")]
        struct SheetView<'a> {
            name: &'a str,
            used_range: Option<String>,
            state: &'static str,
            used_cells: u64,
            formulas: u64,
            errors: u64,
            empty_cells: u64,
            comments: u64,
            charts: u64,
            tables: u64,
            xml_size: u64,
        }
        let sheets = js_sys::Object::new();
        for sheet in &self.described()?.stats.sheets {
            let view = SheetView {
                name: &sheet.name,
                used_range: sheet.end_of_sheet.as_ref().map(|end| format!("A1:{end}")),
                state: match sheet.state {
                    SheetState::Visible => "visible",
                    SheetState::Hidden => "hidden",
                    SheetState::VeryHidden => "very_hidden",
                },
                used_cells: sheet.used_cells,
                formulas: sheet.formulas,
                errors: sheet.errors,
                empty_cells: sheet.empty_cells,
                comments: sheet.comments,
                charts: sheet.charts,
                tables: sheet.tables,
                xml_size: sheet.xml_size,
            };
            js_sys::Reflect::set(&sheets, &JsValue::from_str(&sheet.name), &to_js(&view)?)?;
        }
        Ok(sheets.into())
    }

    /// The defined names and what each refers to, as the file writes it (`{"Rate": "Data!$B$2"}`).
    ///
    /// # Errors
    ///
    /// A message when the result cannot be built.
    #[wasm_bindgen(getter, js_name = definedNames)]
    pub fn defined_names(&self) -> Result<JsValue, JsValue> {
        to_js(&self.described()?.stats.defined_names)
    }

    /// The statistics of the workbook, as `sheetrs stats` reports them.
    #[wasm_bindgen(getter)]
    pub fn stats(&self) -> Result<Stats, JsValue> {
        Ok(Stats {
            inner: Rc::clone(&self.described()?.stats),
        })
    }

    #[wasm_bindgen(getter, js_name = totalSheets)]
    pub fn total_sheets(&self) -> Result<usize, JsValue> {
        Ok(self.described()?.stats.total_sheets)
    }

    #[wasm_bindgen(getter, js_name = totalNamedRanges)]
    pub fn total_named_ranges(&self) -> Result<usize, JsValue> {
        Ok(self.described()?.stats.total_named_ranges)
    }

    /// 0: bytes carry no file size.
    #[wasm_bindgen(getter, js_name = totalFileSize)]
    pub fn total_file_size(&self) -> Result<f64, JsValue> {
        Ok(self.described()?.stats.total_file_size as f64)
    }

    #[wasm_bindgen(getter, js_name = totalUsedCells)]
    pub fn total_used_cells(&self) -> Result<f64, JsValue> {
        Ok(self.described()?.stats.total_used_cells as f64)
    }

    #[wasm_bindgen(getter, js_name = totalFormulas)]
    pub fn total_formulas(&self) -> Result<f64, JsValue> {
        Ok(self.described()?.stats.total_formulas as f64)
    }

    #[wasm_bindgen(getter, js_name = totalErrors)]
    pub fn total_errors(&self) -> Result<f64, JsValue> {
        Ok(self.described()?.stats.total_errors as f64)
    }

    #[wasm_bindgen(getter, js_name = totalEmptyCells)]
    pub fn total_empty_cells(&self) -> Result<f64, JsValue> {
        Ok(self.described()?.stats.total_empty_cells as f64)
    }

    #[wasm_bindgen(getter, js_name = totalComments)]
    pub fn total_comments(&self) -> Result<f64, JsValue> {
        Ok(self.described()?.stats.total_comments as f64)
    }

    #[wasm_bindgen(getter, js_name = totalCharts)]
    pub fn total_charts(&self) -> Result<f64, JsValue> {
        Ok(self.described()?.stats.total_charts as f64)
    }

    #[wasm_bindgen(getter, js_name = totalTables)]
    pub fn total_tables(&self) -> Result<f64, JsValue> {
        Ok(self.described()?.stats.total_tables as f64)
    }

    #[wasm_bindgen(getter, js_name = totalXmlSize)]
    pub fn total_xml_size(&self) -> Result<f64, JsValue> {
        Ok(self.described()?.stats.total_xml_size as f64)
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
    /// The totals and the sheets, as `sheetrs stats --format json` writes them (`totalFileSize` is
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

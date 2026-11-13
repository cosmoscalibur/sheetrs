//! What the operations and the queries give back, as the typed values of the package: `Diagnostic`,
//! `Report`, `ValueRanges`, `Stats` and `Sheet`.
//!
//! Each holds the data of Rust and builds the Python objects when they are asked for; `to_dict()`
//! gives the whole content as the dicts and lists of Python, and `to_json()` the same content as the
//! text of the JSON the command line writes.

use crate::cells::{Block, Grids};
use crate::convert::json_to_py;
use crate::errors::raise;
use pyo3::exceptions::{PyIndexError, PyKeyError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use serde_json::Value;
use sheetrs_compiler::memory::Inspection;
use sheetrs_edit::Skipped;
use sheetrs_io::reader::SheetState;
use sheetrs_protocol::{Diagnostic, Level, MajorDimension, ValueRanges};
use sheetrs_stats::{SheetStats, WorkbookStats};
use std::sync::Arc;

/// The JSON text of a value.
fn json_text(value: &Value) -> PyResult<String> {
    serde_json::to_string(value).map_err(|e| raise("Error", e.to_string()))
}

// ---- Diagnostic -------------------------------------------------------------------------------

/// One thing a run reports besides its result: a `level` (`"error"`, `"warning"`, `"info"` or
/// `"debug"`), a stable `code` that a program can match on, the `message` for a person and the
/// `context` fields the message was made from (`operation`, `reason`, `sheet`, `cell`...).
///
/// The same records go to the `sheetrs` logger of the `logging` module (with the code and the
/// fields in `extra`) and, for the warnings, to `warnings.warn`.
#[pyclass(name = "Diagnostic", frozen, module = "sheetrs")]
#[derive(Clone)]
pub struct PyDiagnostic {
    inner: Diagnostic,
}

impl From<Diagnostic> for PyDiagnostic {
    fn from(inner: Diagnostic) -> Self {
        Self { inner }
    }
}

#[pymethods]
impl PyDiagnostic {
    /// `"error"`, `"warning"`, `"info"` or `"debug"`.
    #[getter]
    fn level(&self) -> &'static str {
        self.inner.level.as_str()
    }

    /// The stable name of what happened (`"RECALC_NOT_SUPPORTED"`).
    #[getter]
    fn code(&self) -> &str {
        &self.inner.code
    }

    #[getter]
    fn message(&self) -> &str {
        &self.inner.message
    }

    /// The fields the message was made from, by name.
    #[getter]
    fn context<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let dict = PyDict::new(py);
        for (key, value) in &self.inner.context {
            dict.set_item(key, value)?;
        }
        Ok(dict)
    }

    /// The record as a dict: `level`, `code`, `message` and `context`.
    fn to_dict<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        Ok(json_to_py(py, &self.inner.to_json())?.into_bound(py))
    }

    /// The record as JSON text.
    fn to_json(&self) -> PyResult<String> {
        json_text(&self.inner.to_json())
    }

    fn __str__(&self) -> String {
        self.inner.to_string()
    }

    fn __repr__(&self) -> String {
        format!(
            "Diagnostic({} {}: {})",
            self.inner.level, self.inner.code, self.inner.message
        )
    }
}

/// The names a `LogRecord` has of its own: a context field of that name cannot go to `extra`.
const RESERVED: &[&str] = &[
    "name",
    "msg",
    "args",
    "levelname",
    "levelno",
    "pathname",
    "filename",
    "module",
    "exc_info",
    "exc_text",
    "stack_info",
    "lineno",
    "funcName",
    "created",
    "msecs",
    "relativeCreated",
    "thread",
    "threadName",
    "processName",
    "process",
    "message",
    "asctime",
    "taskName",
];

/// Sends the diagnostics where Python has them: each to the `sheetrs` logger (`extra` has the `code`
/// and the context fields) and each warning also to `warnings.warn`, as a `RecalculationWarning`
/// for the calculation the engine could not do and a `UserWarning` for any other.
pub(crate) fn emit(py: Python<'_>, diagnostics: &[Diagnostic]) -> PyResult<()> {
    if diagnostics.is_empty() {
        return Ok(());
    }
    let logger = py
        .import("logging")?
        .call_method1("getLogger", ("sheetrs",))?;
    let warnings = py.import("warnings")?;
    for diagnostic in diagnostics {
        let level: u32 = match diagnostic.level {
            Level::Error => 40,
            Level::Warning => 30,
            Level::Debug => 10,
            // Information, and a level this does not know (`Level` may grow).
            _ => 20,
        };
        let extra = PyDict::new(py);
        for (key, value) in &diagnostic.context {
            let key = if RESERVED.contains(&key.as_str()) {
                format!("{key}_")
            } else {
                key.clone()
            };
            extra.set_item(key, value)?;
        }
        extra.set_item("code", &diagnostic.code)?;
        let kwargs = PyDict::new(py);
        kwargs.set_item("extra", extra)?;
        logger.call_method("log", (level, &diagnostic.message), Some(&kwargs))?;
        if diagnostic.level == Level::Warning {
            let class = if diagnostic.code == sheetrs_edit::RECALC_NOT_SUPPORTED {
                py.import("sheetrs._errors")?
                    .getattr("RecalculationWarning")?
            } else {
                py.import("builtins")?.getattr("UserWarning")?
            };
            warnings.call_method1("warn", (&diagnostic.message, class, 1))?;
        }
    }
    Ok(())
}

// ---- Report -----------------------------------------------------------------------------------

/// A cell a `replace` found and left as it was, and why.
#[pyclass(name = "Skipped", frozen, module = "sheetrs")]
#[derive(Clone)]
pub struct PySkipped {
    /// The sheet.
    #[pyo3(get)]
    sheet: String,
    /// The cell in A1 notation.
    #[pyo3(get)]
    cell: String,
    /// Why it was left.
    #[pyo3(get)]
    reason: String,
}

#[pymethods]
impl PySkipped {
    fn to_dict<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let dict = PyDict::new(py);
        dict.set_item("sheet", &self.sheet)?;
        dict.set_item("cell", &self.cell)?;
        dict.set_item("reason", &self.reason)?;
        Ok(dict)
    }

    fn __repr__(&self) -> String {
        format!("Skipped({}!{}: {})", self.sheet, self.cell, self.reason)
    }
}

impl From<Skipped> for PySkipped {
    fn from(s: Skipped) -> Self {
        Self {
            sheet: s.sheet,
            cell: s.cell,
            reason: s.reason,
        }
    }
}

/// What a run did besides writing the workbook: the cells a `replace` changed (`replaced`), the
/// cells written (`written`), the cells it found and left (`skipped`) and the `diagnostics`.
#[pyclass(name = "Report", frozen, module = "sheetrs")]
#[derive(Clone, Default)]
pub struct PyReport {
    #[pyo3(get)]
    replaced: usize,
    #[pyo3(get)]
    written: usize,
    skipped: Vec<PySkipped>,
    diagnostics: Vec<Diagnostic>,
}

impl PyReport {
    /// The report of a run of the editor.
    pub(crate) fn of(report: &sheetrs_edit::Report) -> Self {
        Self {
            replaced: report.replaced,
            written: report.written,
            skipped: report.skipped.iter().cloned().map(Into::into).collect(),
            diagnostics: report.diagnostics.clone(),
        }
    }

    fn json(&self) -> Value {
        serde_json::json!({
            "replaced": self.replaced,
            "written": self.written,
            "skipped": self.skipped.iter().map(|s| serde_json::json!({
                "sheet": s.sheet, "cell": s.cell, "reason": s.reason,
            })).collect::<Vec<_>>(),
            "diagnostics": self.diagnostics.iter().map(Diagnostic::to_json).collect::<Vec<_>>(),
        })
    }
}

#[pymethods]
impl PyReport {
    /// The cells a `replace` found and left as they were.
    #[getter]
    fn skipped(&self) -> Vec<PySkipped> {
        self.skipped.clone()
    }

    /// What the run reported besides its result, as `Diagnostic` records.
    #[getter]
    fn diagnostics(&self) -> Vec<PyDiagnostic> {
        self.diagnostics.iter().cloned().map(Into::into).collect()
    }

    fn to_dict<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        Ok(json_to_py(py, &self.json())?.into_bound(py))
    }

    fn to_json(&self) -> PyResult<String> {
        json_text(&self.json())
    }

    fn __repr__(&self) -> String {
        format!(
            "Report(replaced={}, written={}, skipped={}, diagnostics={})",
            self.replaced,
            self.written,
            self.skipped.len(),
            self.diagnostics.len()
        )
    }
}

// ---- ValueRanges ------------------------------------------------------------------------------

/// The value ranges of a calculation: a mapping from the name of each one (the range it was asked
/// with, or the name of the declared output) to a `Block` with its values.
///
/// `value_ranges["Summary!B2:B5"]` is a `Block`; it has `keys()`, `values()`, `items()`, `len` and
/// membership like a `dict`. `to_dict()` gives `{name: list[list]}` and `to_json()` the JSON answer
/// of `sheetc calculate`.
#[pyclass(name = "ValueRanges", frozen, module = "sheetrs")]
pub struct PyValueRanges {
    value_ranges: ValueRanges,
    entries: Vec<(String, String, Value)>,
    major: &'static str,
}

impl PyValueRanges {
    /// The value ranges `asked` for (`None`: the outputs the workbook declares or, with none, its
    /// formula cells).
    pub(crate) fn new(
        value_ranges: ValueRanges,
        asked: Option<&[String]>,
        major: MajorDimension,
    ) -> Self {
        let ranges: Vec<String> = value_ranges
            .value_ranges
            .iter()
            .map(|r| r.range.clone())
            .collect();
        let entries = value_ranges
            .to_entries(asked)
            .into_iter()
            .enumerate()
            .map(|(i, (name, rows))| {
                let range = ranges.get(i).cloned().unwrap_or_else(|| name.clone());
                (name, range, rows)
            })
            .collect();
        Self {
            value_ranges,
            entries,
            major: if major == MajorDimension::Columns {
                "columns"
            } else {
                "rows"
            },
        }
    }

    fn block(
        py: Python<'_>,
        range: &str,
        major: &'static str,
        rows: &Value,
    ) -> PyResult<Py<Block>> {
        let list = json_to_py(py, rows)?
            .cast_bound::<PyList>(py)?
            .clone()
            .unbind();
        Py::new(
            py,
            Block::new(
                range.to_owned(),
                major,
                Grids {
                    values: Some(list),
                    formulas: None,
                    number_formats: None,
                    displays: None,
                    comments: None,
                },
            ),
        )
    }
}

#[pymethods]
impl PyValueRanges {
    fn __len__(&self) -> usize {
        self.entries.len()
    }

    fn __contains__(&self, key: &str) -> bool {
        self.entries.iter().any(|(name, _, _)| name == key)
    }

    fn __getitem__(&self, py: Python<'_>, key: &str) -> PyResult<Py<Block>> {
        let (_, range, rows) = self
            .entries
            .iter()
            .find(|(name, _, _)| name == key)
            .ok_or_else(|| PyKeyError::new_err(key.to_owned()))?;
        Self::block(py, range, self.major, rows)
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let keys = PyList::new(py, self.entries.iter().map(|(name, _, _)| name))?;
        Ok(keys.try_iter()?.into_any().unbind())
    }

    /// The names of the value ranges, in order.
    fn keys(&self) -> Vec<String> {
        self.entries
            .iter()
            .map(|(name, _, _)| name.clone())
            .collect()
    }

    /// The `Block` of each value range, in order.
    fn values(&self, py: Python<'_>) -> PyResult<Vec<Py<Block>>> {
        self.entries
            .iter()
            .map(|(_, range, rows)| Self::block(py, range, self.major, rows))
            .collect()
    }

    /// `(name, Block)` pairs, in order.
    fn items(&self, py: Python<'_>) -> PyResult<Vec<(String, Py<Block>)>> {
        self.entries
            .iter()
            .map(|(name, range, rows)| {
                Ok((name.clone(), Self::block(py, range, self.major, rows)?))
            })
            .collect()
    }

    /// `{name: list[list]}`: the values of each value range as the lists of rows Python has.
    fn to_dict<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let dict = PyDict::new(py);
        for (name, _, rows) in &self.entries {
            dict.set_item(name, json_to_py(py, rows)?)?;
        }
        Ok(dict)
    }

    /// The JSON answer of the protocol (`{"valueRanges": [...]}`) as text.
    fn to_json(&self) -> PyResult<String> {
        json_text(&self.value_ranges.to_json())
    }

    fn __repr__(&self) -> String {
        format!("ValueRanges({})", self.keys().join(", "))
    }
}

// ---- Stats and Sheet -------------------------------------------------------------------------

/// A sheet of the workbook, with what the read at open found out about it.
#[pyclass(name = "Sheet", frozen, module = "sheetrs")]
#[derive(Clone)]
pub struct PySheet {
    inner: SheetStats,
}

#[pymethods]
impl PySheet {
    #[getter]
    fn name(&self) -> &str {
        &self.inner.name
    }

    /// The area from `A1` to the bottom right corner of the cells that hold a value or a formula
    /// (`"A1:GV5439"`); `None` for a sheet with none.
    #[getter]
    fn used_range(&self) -> Option<String> {
        self.inner
            .end_of_sheet
            .as_ref()
            .map(|end| format!("A1:{end}"))
    }

    /// `"visible"`, `"hidden"` or `"very_hidden"`.
    #[getter]
    fn state(&self) -> &'static str {
        match self.inner.state {
            SheetState::Visible => "visible",
            SheetState::Hidden => "hidden",
            SheetState::VeryHidden => "very_hidden",
        }
    }

    /// Cells that hold a value or a formula.
    #[getter]
    fn used_cells(&self) -> u64 {
        self.inner.used_cells
    }

    #[getter]
    fn formulas(&self) -> u64 {
        self.inner.formulas
    }

    /// Cells whose value (the cached result of a formula) is an error.
    #[getter]
    fn errors(&self) -> u64 {
        self.inner.errors
    }

    /// Cells without value or formula inside the used range.
    #[getter]
    fn empty_cells(&self) -> u64 {
        self.inner.empty_cells
    }

    #[getter]
    fn comments(&self) -> u64 {
        self.inner.comments
    }

    #[getter]
    fn charts(&self) -> u64 {
        self.inner.charts
    }

    /// Table objects of an XLSX, named database ranges of an ODS.
    #[getter]
    fn tables(&self) -> u64 {
        self.inner.tables
    }

    /// Uncompressed size of the XML that describes the sheet.
    #[getter]
    fn xml_size(&self) -> u64 {
        self.inner.xml_size
    }

    /// The statistics of the sheet as `sheetrs stats --format json` gives them.
    fn to_dict<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let value = serde_json::to_value(&self.inner).map_err(|e| raise("Error", e.to_string()))?;
        Ok(json_to_py(py, &value)?.into_bound(py))
    }

    fn to_json(&self) -> PyResult<String> {
        json_text(&serde_json::to_value(&self.inner).map_err(|e| raise("Error", e.to_string()))?)
    }

    fn __repr__(&self) -> String {
        format!("Sheet({:?})", self.inner.name)
    }
}

/// The sheets of a workbook, by name or by position, in the order of the workbook: `book.sheets["Data"]`,
/// `book.sheets[0]`, `for sheet in book.sheets`, `len(book.sheets)`.
#[pyclass(name = "Sheets", frozen, module = "sheetrs")]
pub struct PySheets {
    sheets: Vec<PySheet>,
}

#[pymethods]
impl PySheets {
    fn __len__(&self) -> usize {
        self.sheets.len()
    }

    fn __contains__(&self, name: &str) -> bool {
        self.sheets.iter().any(|s| s.inner.name == name)
    }

    fn __getitem__(&self, key: &Bound<'_, PyAny>) -> PyResult<PySheet> {
        if let Ok(name) = key.extract::<String>() {
            return self
                .sheets
                .iter()
                .find(|s| s.inner.name == name)
                .cloned()
                .ok_or_else(|| PyKeyError::new_err(name));
        }
        let index: isize = key
            .extract()
            .map_err(|_| PyKeyError::new_err("a sheet is looked up by its name or its position"))?;
        let len = isize::try_from(self.sheets.len()).unwrap_or(isize::MAX);
        let at = if index < 0 { index + len } else { index };
        usize::try_from(at)
            .ok()
            .and_then(|i| self.sheets.get(i))
            .cloned()
            .ok_or_else(|| PyIndexError::new_err("sheet index out of range"))
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let list = PyList::new(py, self.sheets.iter().cloned())?;
        Ok(list.try_iter()?.into_any().unbind())
    }

    /// The names of the sheets, in order.
    fn names(&self) -> Vec<String> {
        self.sheets.iter().map(|s| s.inner.name.clone()).collect()
    }

    fn __repr__(&self) -> String {
        format!("Sheets({})", self.names().join(", "))
    }
}

impl PySheets {
    pub(crate) fn of(stats: &WorkbookStats) -> Self {
        Self {
            sheets: stats
                .sheets
                .iter()
                .cloned()
                .map(|inner| PySheet { inner })
                .collect(),
        }
    }
}

/// The statistics of the workbook: the eleven totals and the `sheets`, as `sheetrs stats` reports
/// them. `to_dict()` is its `--format json`.
#[pyclass(name = "Stats", frozen, module = "sheetrs")]
#[derive(Clone)]
pub struct PyStats {
    inner: Arc<WorkbookStats>,
}

impl PyStats {
    pub(crate) fn of(inner: Arc<WorkbookStats>) -> Self {
        Self { inner }
    }
}

#[pymethods]
impl PyStats {
    #[getter]
    fn total_sheets(&self) -> usize {
        self.inner.total_sheets
    }

    #[getter]
    fn total_named_ranges(&self) -> usize {
        self.inner.total_named_ranges
    }

    /// The size of the file in bytes; 0 for a workbook given as bytes or as a stream.
    #[getter]
    fn total_file_size(&self) -> u64 {
        self.inner.total_file_size
    }

    #[getter]
    fn total_used_cells(&self) -> u64 {
        self.inner.total_used_cells
    }

    #[getter]
    fn total_formulas(&self) -> u64 {
        self.inner.total_formulas
    }

    #[getter]
    fn total_errors(&self) -> u64 {
        self.inner.total_errors
    }

    #[getter]
    fn total_empty_cells(&self) -> u64 {
        self.inner.total_empty_cells
    }

    #[getter]
    fn total_comments(&self) -> u64 {
        self.inner.total_comments
    }

    #[getter]
    fn total_charts(&self) -> u64 {
        self.inner.total_charts
    }

    #[getter]
    fn total_tables(&self) -> u64 {
        self.inner.total_tables
    }

    #[getter]
    fn total_xml_size(&self) -> u64 {
        self.inner.total_xml_size
    }

    /// The sheets, by name or by position.
    #[getter]
    fn sheets(&self) -> PySheets {
        PySheets::of(&self.inner)
    }

    fn to_dict<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let value =
            serde_json::to_value(&*self.inner).map_err(|e| raise("Error", e.to_string()))?;
        Ok(json_to_py(py, &value)?.into_bound(py))
    }

    fn to_json(&self) -> PyResult<String> {
        json_text(&serde_json::to_value(&*self.inner).map_err(|e| raise("Error", e.to_string()))?)
    }

    fn __repr__(&self) -> String {
        format!(
            "Stats({} sheets, {} cells, {} formulas)",
            self.inner.total_sheets, self.inner.total_used_cells, self.inner.total_formulas
        )
    }
}

// ---- Inspection ------------------------------------------------------------------------------

/// A function a workbook uses.
#[pyclass(name = "FunctionUse", frozen, module = "sheetrs")]
#[derive(Clone)]
pub struct PyFunctionUse {
    #[pyo3(get)]
    name: String,
    /// The number of calls.
    #[pyo3(get)]
    calls: usize,
    /// Whether the engine has it.
    #[pyo3(get)]
    implemented: bool,
}

#[pymethods]
impl PyFunctionUse {
    fn __repr__(&self) -> String {
        format!("FunctionUse({}, {} calls)", self.name, self.calls)
    }
}

/// What `sheetc inspect` says about a workbook.
#[pyclass(name = "Inspection", frozen, module = "sheetrs")]
pub struct PyInspection {
    inner: Inspection,
    functions: Vec<PyFunctionUse>,
}

impl PyInspection {
    pub(crate) fn of(inner: Inspection) -> Self {
        let functions = inner
            .functions
            .iter()
            .map(|f| PyFunctionUse {
                name: f.name.clone(),
                calls: f.calls,
                implemented: f.implemented,
            })
            .collect();
        Self { inner, functions }
    }
}

#[pymethods]
impl PyInspection {
    /// The names of the sheets.
    #[getter]
    fn sheets(&self) -> Vec<String> {
        self.inner.sheets.clone()
    }

    /// The null date of the workbook (`"1899-12-30"`, `"1900-01-01"` or `"1904-01-01"`).
    #[getter]
    fn date_system(&self) -> &str {
        &self.inner.date_system
    }

    /// Cells with a value or a formula.
    #[getter]
    fn cells(&self) -> usize {
        self.inner.cells
    }

    #[getter]
    fn formulas(&self) -> usize {
        self.inner.formulas
    }

    /// Groups of cells that read one another.
    #[getter]
    fn circular_groups(&self) -> usize {
        self.inner.circular_groups
    }

    /// `INDIRECT` or `OFFSET` are used: the whole workbook is reachable.
    #[getter]
    fn run_time_references(&self) -> bool {
        self.inner.run_time_references
    }

    /// Functions the workbook uses that the engine does not have (and `LibreOffice` does not know
    /// either).
    #[getter]
    fn unsupported_functions(&self) -> Vec<String> {
        self.inner.unsupported_functions.clone()
    }

    /// The functions used, by number of calls, most used first.
    #[getter]
    fn functions(&self) -> Vec<PyFunctionUse> {
        self.functions.clone()
    }

    /// What the compiler says about the workbook while it reads it.
    #[getter]
    fn diagnostics(&self) -> Vec<String> {
        self.inner.diagnostics.clone()
    }

    fn to_dict<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let value = serde_json::to_value(&self.inner).map_err(|e| raise("Error", e.to_string()))?;
        Ok(json_to_py(py, &value)?.into_bound(py))
    }

    fn to_json(&self) -> PyResult<String> {
        json_text(&serde_json::to_value(&self.inner).map_err(|e| raise("Error", e.to_string()))?)
    }

    fn __repr__(&self) -> String {
        format!(
            "Inspection({} sheets, {} cells, {} formulas)",
            self.inner.sheets.len(),
            self.inner.cells,
            self.inner.formulas
        )
    }
}

//! The linter and the statistics of a `Book`.

use crate::errors::to_py_err;
use pyo3::exceptions::{PyIndexError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use sheetrs_io::Book;
use sheetrs_linter::report::ViolationRecord;
use sheetrs_linter::rules::registry;
use sheetrs_linter::{BookLint, LintReport, Linter, LinterConfig, OutputFormat};
use std::path::Path;
use std::sync::OnceLock;

/// One finding of the linter.
#[pyclass(name = "Violation", frozen, module = "sheetrs")]
#[derive(Clone)]
pub struct PyViolation {
    #[pyo3(get)]
    rule_id: String,
    #[pyo3(get)]
    rule_name: String,
    /// `"error"`, `"warning"` or `"info"`.
    #[pyo3(get)]
    severity: &'static str,
    #[pyo3(get)]
    message: String,
    /// `"book"`, `"sheet"` or `"cell"`.
    #[pyo3(get)]
    scope: &'static str,
    /// Name of the sheet, `None` for a book-level violation.
    #[pyo3(get)]
    sheet: Option<String>,
    /// Cell in A1 notation, `None` unless the scope is `"cell"`.
    #[pyo3(get)]
    cell: Option<String>,
    /// 0-based row of the cell.
    #[pyo3(get)]
    row: Option<u32>,
    /// 0-based column of the cell.
    #[pyo3(get)]
    col: Option<u32>,
}

#[pymethods]
impl PyViolation {
    fn to_dict<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let dict = PyDict::new(py);
        dict.set_item("rule_id", &self.rule_id)?;
        dict.set_item("rule_name", &self.rule_name)?;
        dict.set_item("severity", self.severity)?;
        dict.set_item("message", &self.message)?;
        dict.set_item("scope", self.scope)?;
        dict.set_item("sheet", &self.sheet)?;
        dict.set_item("cell", &self.cell)?;
        dict.set_item("row", self.row)?;
        dict.set_item("col", self.col)?;
        Ok(dict)
    }

    fn __repr__(&self) -> String {
        let at = match (&self.sheet, &self.cell) {
            (Some(sheet), Some(cell)) => format!("{sheet}!{cell}"),
            (Some(sheet), None) => sheet.clone(),
            _ => "book".to_owned(),
        };
        format!(
            "Violation({} {} at {at}: {})",
            self.rule_id, self.severity, self.message
        )
    }

    fn __str__(&self) -> String {
        self.__repr__()
    }
}

impl From<ViolationRecord> for PyViolation {
    fn from(record: ViolationRecord) -> Self {
        Self {
            rule_id: record.rule_id,
            rule_name: record.rule_name,
            severity: record.severity,
            message: record.message,
            scope: record.scope,
            sheet: record.sheet,
            cell: record.cell,
            row: record.row,
            col: record.col,
        }
    }
}

/// The linter with the configuration of `config_path` (the default rules without it). A rule,
/// category or parameter the linter does not know is a `ValueError`.
pub(crate) fn linter(config_path: Option<&Path>) -> PyResult<Linter> {
    let config = match config_path {
        Some(path) => LinterConfig::from_file(path),
        None => Ok(LinterConfig::default()),
    }
    .map_err(|e| PyValueError::new_err(format!("invalid configuration: {e:#}")))?;
    Ok(Linter::with_config(config))
}

/// The result of a lint run: the violations, ordered by scope (book, sheet, cell) and rule, and the
/// means to write them as the three documents of `sheetrs lint`.
///
/// It is a sequence of `Violation` (`len`, iteration, indexing, slices); the objects are made when they
/// are asked for. `to_text()` is the report for a person, `to_json()` the JSON document and
/// `to_sarif()` the SARIF 2.1.0 log, all as text; `to_dict()` is the JSON document as the dicts and lists
/// of Python. The `file` of the three is the name the document gives the workbook.
#[pyclass(name = "LintReport", frozen, module = "sheetrs")]
pub struct PyLintReport {
    report: LintReport,
    records: OnceLock<Vec<ViolationRecord>>,
}

impl PyLintReport {
    fn records(&self) -> &[ViolationRecord] {
        self.records.get_or_init(|| self.report.records())
    }

    fn render(&self, py: Python<'_>, format: OutputFormat, file: Option<&str>) -> PyResult<String> {
        py.detach(|| {
            let mut out = Vec::new();
            self.report
                .write(&mut out, format, file)
                .map_err(to_py_err)?;
            String::from_utf8(out).map_err(to_py_err)
        })
    }
}

#[pymethods]
impl PyLintReport {
    fn __len__(&self) -> usize {
        self.report.violations().len()
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let list = PyList::new(py, self.violations())?;
        Ok(list.try_iter()?.into_any().unbind())
    }

    fn __getitem__<'py>(
        &self,
        py: Python<'py>,
        index: &Bound<'py, PyAny>,
    ) -> PyResult<Bound<'py, PyAny>> {
        PyList::new(py, self.violations())?
            .into_any()
            .get_item(index)
            .map_err(|_| PyIndexError::new_err("violation index out of range"))
    }

    /// The violations, as a list.
    #[getter]
    fn violations(&self) -> Vec<PyViolation> {
        self.records()
            .iter()
            .cloned()
            .map(PyViolation::from)
            .collect()
    }

    /// The report for a person, as plain text.
    #[pyo3(signature = (file=None))]
    fn to_text(&self, py: Python<'_>, file: Option<&str>) -> PyResult<String> {
        self.render(py, OutputFormat::Human, file)
    }

    /// The JSON document of `sheetrs lint --format json` (`file`, `summary`, `violations`), as text.
    #[pyo3(signature = (file=None))]
    fn to_json(&self, py: Python<'_>, file: Option<&str>) -> PyResult<String> {
        self.render(py, OutputFormat::Json, file)
    }

    /// The SARIF 2.1.0 log of `sheetrs lint --format sarif`, as text.
    #[pyo3(signature = (file=None))]
    fn to_sarif(&self, py: Python<'_>, file: Option<&str>) -> PyResult<String> {
        self.render(py, OutputFormat::Sarif, file)
    }

    /// The JSON document of `to_json()` as dicts and lists.
    #[pyo3(signature = (file=None))]
    fn to_dict<'py>(&self, py: Python<'py>, file: Option<&str>) -> PyResult<Bound<'py, PyAny>> {
        let text = self.render(py, OutputFormat::Json, file)?;
        py.import("json")?.call_method1("loads", (text,))
    }

    fn __repr__(&self) -> String {
        format!("LintReport({} violations)", self.report.violations().len())
    }
}

/// Lints the workbook of `source`.
pub(crate) fn lint(
    py: Python<'_>,
    source: &Book,
    config_path: Option<&Path>,
) -> PyResult<PyLintReport> {
    let linter = linter(config_path)?;
    let report = py.detach(|| source.lint(&linter)).map_err(to_py_err)?;
    Ok(PyLintReport {
        report,
        records: OnceLock::new(),
    })
}

/// The rules, ordered by id, for the configuration of `config_path`.
pub(crate) fn rules<'py>(
    py: Python<'py>,
    config_path: Option<&Path>,
) -> PyResult<Bound<'py, PyList>> {
    let linter = linter(config_path)?;
    let config = linter.config();
    let list = PyList::empty(py);
    for meta in registry::rule_metadata(config) {
        let item = PyDict::new(py);
        item.set_item("enabled", registry::is_rule_active(meta.id, config))?;
        item.set_item("id", meta.id.as_str())?;
        item.set_item("name", meta.name)?;
        item.set_item("category", meta.category.as_str())?;
        item.set_item("default", meta.is_default)?;
        list.append(item)?;
    }
    Ok(list)
}

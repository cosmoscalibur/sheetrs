//! Python bindings for the SheetRS suite.
//!
//! The extension module is `sheetrs._sheetrs`; the `sheetrs` package re-exports it. It exposes
//! the three tools as functions that take a path (`str`/`os.PathLike`) or the bytes of a workbook:
//!
//! - `lint`: `sheetlint`, returning `Violation` objects.
//! - `stats`: `sheetstats`, returning a `dict`.
//! - `edit`: `sheetedit`, writing a modified copy of a workbook.
//!
//! The functions release the GIL while they read and analyze a workbook.

use pyo3::exceptions::{PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList};
use sheetrs_core::reader::{self, Workbook};
use sheetrs_core::report::{ViolationRecord, violation_records};
use sheetrs_core::rules::registry;
use sheetrs_core::writer::{self, ValueScope, WorkbookModifications};
use sheetrs_core::{Linter, LinterConfig};
use sheetrs_stats::stats::WorkbookStats;
use std::collections::HashSet;
use std::io::Cursor;
use std::path::PathBuf;

pyo3::create_exception!(
    _sheetrs,
    SheetrsError,
    pyo3::exceptions::PyException,
    "A workbook could not be read, analyzed or written."
);

/// A workbook given by path or as bytes.
enum Source {
    Path(PathBuf),
    Bytes(Vec<u8>, String),
}

impl Source {
    /// `source` is a path (`str`, `os.PathLike`) or `bytes`; bytes need `extension`
    /// (`xlsx`, `xlsm` or `ods`) because the format is not read from the content.
    fn from_py(source: &Bound<'_, PyAny>, extension: Option<&str>) -> PyResult<Self> {
        if let Ok(bytes) = source.cast::<PyBytes>() {
            let Some(extension) = extension else {
                return Err(PyValueError::new_err(
                    "`extension` ('xlsx', 'xlsm' or 'ods') is required to read a workbook from bytes",
                ));
            };
            return Ok(Self::Bytes(bytes.as_bytes().to_vec(), extension.to_owned()));
        }
        source
            .extract::<PathBuf>()
            .map(Self::Path)
            .map_err(|_| PyTypeError::new_err("a workbook is a path (str or os.PathLike) or bytes"))
    }

    fn read(&self) -> anyhow::Result<Workbook> {
        match self {
            Self::Path(path) => reader::read_workbook(path),
            Self::Bytes(bytes, extension) => {
                reader::read_workbook_from_reader(Cursor::new(bytes), Some(extension))
            }
        }
    }

    fn stats(&self) -> anyhow::Result<WorkbookStats> {
        match self {
            Self::Path(path) => WorkbookStats::from_path(path),
            Self::Bytes(bytes, extension) => {
                WorkbookStats::from_reader(Cursor::new(bytes), Some(extension))
            }
        }
    }
}

/// An I/O failure becomes `OSError` (`FileNotFoundError`, …); anything else, `SheetrsError`.
fn to_py_err(error: anyhow::Error) -> PyErr {
    match error.downcast::<std::io::Error>() {
        Ok(io) => io.into(),
        Err(error) => SheetrsError::new_err(format!("{error:#}")),
    }
}

/// One finding of the linter.
#[pyclass(name = "Violation", frozen, module = "sheetrs")]
#[derive(Clone)]
struct PyViolation {
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

/// Lint a workbook.
///
/// `source` is a path or the bytes of an XLSX/XLSM/ODS file (then `extension` is required).
/// The configuration is the TOML text of a `sheetlint.toml` (`config`) or the path of one
/// (`config_path`), not both; without either, the default rules run.
/// Returns the violations ordered by scope (book, then sheet, then cell) and rule.
///
/// Raises `ValueError` for an invalid configuration, `OSError` when a file cannot be read and
/// `SheetrsError` when the workbook cannot be parsed.
#[pyfunction]
#[pyo3(signature = (source, *, config=None, config_path=None, extension=None))]
fn lint(
    py: Python<'_>,
    source: &Bound<'_, PyAny>,
    config: Option<&str>,
    config_path: Option<PathBuf>,
    extension: Option<&str>,
) -> PyResult<Vec<PyViolation>> {
    let source = Source::from_py(source, extension)?;
    let config = match (config, config_path) {
        (Some(_), Some(_)) => {
            return Err(PyValueError::new_err(
                "give `config` or `config_path`, not both",
            ));
        }
        (Some(text), None) => LinterConfig::from_toml(text),
        (None, Some(path)) => LinterConfig::from_file(path),
        (None, None) => Ok(LinterConfig::default()),
    }
    .map_err(|e| PyValueError::new_err(format!("invalid configuration: {e:#}")))?;

    py.detach(|| {
        let workbook = source.read().map_err(to_py_err)?;
        let linter = Linter::with_config(config);
        let violations = linter.lint_workbook(&workbook).map_err(to_py_err)?;
        Ok(violation_records(&violations, &workbook, linter.config())
            .into_iter()
            .map(PyViolation::from)
            .collect())
    })
}

/// The lint rules: a list of dicts with `id`, `name`, `category` and `default` (whether the rule
/// runs without configuration).
#[pyfunction]
fn rules(py: Python<'_>) -> PyResult<Bound<'_, PyList>> {
    let list = PyList::empty(py);
    for meta in registry::get_all_rule_metadata(&LinterConfig::default()) {
        let item = PyDict::new(py);
        item.set_item("id", meta.id.as_str())?;
        item.set_item("name", meta.name)?;
        item.set_item("category", meta.category.as_str())?;
        item.set_item("default", meta.is_default)?;
        list.append(item)?;
    }
    Ok(list)
}

/// Statistics of a workbook, as `sheetstats --format json` reports them: totals plus a
/// `sheets` list with one dict per sheet.
///
/// The file size is 0 for a workbook given as bytes.
#[pyfunction]
#[pyo3(signature = (source, *, extension=None))]
fn stats<'py>(
    py: Python<'py>,
    source: &Bound<'py, PyAny>,
    extension: Option<&str>,
) -> PyResult<Bound<'py, PyAny>> {
    let source = Source::from_py(source, extension)?;
    let stats = py.detach(|| source.stats().map_err(to_py_err))?;
    let json = serde_json::to_string(&stats).map_err(|e| SheetrsError::new_err(e.to_string()))?;
    py.import("json")?.call_method1("loads", (json,))
}

/// Write a modified copy of the workbook at `input` to `output` (XLSX or ODS, by extension).
///
/// - `remove_sheets`, `remove_ranges`: names to remove.
/// - `convert_to_values`: scopes (`"Sheet"`, `"Sheet!B:D"`, `"Sheet!3:10"`) whose formulas are
///   replaced by their cached values.
/// - `delete_rows` (`"Sheet!3:10"`), `delete_cols` (`"Sheet!B:D"`): in the coordinates of the
///   input; references to them become `#REF!`.
/// - `recalc_on_open`: the application that opens the output recalculates every formula.
///
/// Raises `ValueError` for a malformed scope and `SheetrsError` when the operation fails (the
/// output is not left behind).
#[pyfunction]
#[pyo3(signature = (
    input,
    output,
    *,
    remove_sheets=None,
    remove_ranges=None,
    convert_to_values=None,
    delete_rows=None,
    delete_cols=None,
    recalc_on_open=false,
))]
#[allow(clippy::too_many_arguments)]
fn edit(
    py: Python<'_>,
    input: PathBuf,
    output: PathBuf,
    remove_sheets: Option<Vec<String>>,
    remove_ranges: Option<Vec<String>>,
    convert_to_values: Option<Vec<String>>,
    delete_rows: Option<Vec<String>>,
    delete_cols: Option<Vec<String>>,
    recalc_on_open: bool,
) -> PyResult<()> {
    fn scopes(specs: Option<Vec<String>>) -> PyResult<Vec<ValueScope>> {
        specs
            .unwrap_or_default()
            .iter()
            .map(|spec| {
                spec.parse::<ValueScope>()
                    .map_err(|e| PyValueError::new_err(format!("invalid scope '{spec}': {e:#}")))
            })
            .collect()
    }
    let modifications = WorkbookModifications {
        remove_sheets: remove_sheets.map(|names| names.into_iter().collect::<HashSet<_>>()),
        remove_named_ranges: remove_ranges.map(|names| names.into_iter().collect::<HashSet<_>>()),
        convert_to_values: scopes(convert_to_values)?,
        delete_rows: scopes(delete_rows)?,
        delete_columns: scopes(delete_cols)?,
        recalculate_on_open: recalc_on_open,
    };
    py.detach(|| writer::modify_workbook(&input, &output, &modifications).map_err(to_py_err))
}

#[pymodule(gil_used = false)]
fn _sheetrs(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add("SheetrsError", module.py().get_type::<SheetrsError>())?;
    module.add_class::<PyViolation>()?;
    module.add_function(wrap_pyfunction!(lint, module)?)?;
    module.add_function(wrap_pyfunction!(rules, module)?)?;
    module.add_function(wrap_pyfunction!(stats, module)?)?;
    module.add_function(wrap_pyfunction!(edit, module)?)?;
    module.add("__version__", env!("CARGO_PKG_VERSION"))?;
    Ok(())
}

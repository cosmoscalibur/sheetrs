//! Python bindings for the SheetRS suite.
//!
//! The extension module is `sheetrs._sheetrs`; the `sheetrs` package re-exports it, flat. A
//! workbook is a path (`str`/`os.PathLike`) or its bytes. The functions release the GIL while
//! they read and analyze a workbook.
//!
//! - `Linter`: a configuration validated when it is built, then `lint` and `rules` (the rules
//!   exist only as part of the linter, so their catalog is next to the configuration it reflects).
//! - `stats`: the statistics of a workbook.
//! - `open` gives a `Workbook` to edit: operations are recorded, each addressing the workbook as
//!   the previous ones leave it, and `save` runs them over the file in as few passes as it can.
//!   Editing has no eager mode: a workbook is streamed, never loaded, so there is no model to
//!   apply an operation to before `save`.

use pyo3::exceptions::{PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList, PyTuple};
use sheetrs_core::reader::{self, Workbook};
use sheetrs_edit::{ColumnRange, EditPlan, Operation, RowRange, ValueScope};
use sheetrs_linter::report::{ViolationRecord, violation_records};
use sheetrs_linter::rules::registry;
use sheetrs_linter::{Linter, LinterConfig};
use sheetrs_stats::WorkbookStats;
use std::io::Cursor;
use std::path::PathBuf;
use std::sync::Arc;

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

/// The linter, with a configuration that was validated when it was built.
///
/// The configuration is the TOML text of a `sheetlint.toml` (`config`) or the path of one
/// (`config_path`), not both; without either, the default rules run. A rule, category or
/// parameter the linter does not know is an error (`ValueError`), not something ignored.
/// Build it once to lint many workbooks.
#[pyclass(name = "Linter", frozen, module = "sheetrs")]
struct PyLinter {
    linter: Linter,
}

#[pymethods]
impl PyLinter {
    #[new]
    #[pyo3(signature = (*, config=None, config_path=None))]
    fn new(config: Option<&str>, config_path: Option<PathBuf>) -> PyResult<Self> {
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
        Ok(Self {
            linter: Linter::with_config(config),
        })
    }

    /// Lint a workbook, given by path or as the bytes of an XLSX/XLSM/ODS file (then `extension`
    /// is required). Returns the violations ordered by scope (book, sheet, cell) and rule.
    ///
    /// Raises `OSError` when a file cannot be read and `SheetrsError` when the workbook cannot be
    /// parsed.
    #[pyo3(signature = (source, *, extension=None))]
    fn lint(
        &self,
        py: Python<'_>,
        source: &Bound<'_, PyAny>,
        extension: Option<&str>,
    ) -> PyResult<Vec<PyViolation>> {
        let source = Source::from_py(source, extension)?;
        py.detach(|| {
            let workbook = source.read().map_err(to_py_err)?;
            let violations = self.linter.lint_workbook(&workbook).map_err(to_py_err)?;
            Ok(
                violation_records(&violations, &workbook, self.linter.config())
                    .into_iter()
                    .map(PyViolation::from)
                    .collect(),
            )
        })
    }

    /// The rules, ordered by id: a list of dicts with `id`, `name`, `category`, `default` (runs
    /// without configuration) and `enabled` (runs with this linter's configuration).
    ///
    /// The ids and categories are what `disabled_rules` and `enabled_rules` accept.
    fn rules<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let config = self.linter.config();
        let list = PyList::empty(py);
        for meta in registry::get_all_rule_metadata(config) {
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
}

/// Statistics of a workbook, as `sheetstats --format json` reports them: totals plus a
/// `sheets` list with one dict per sheet.
///
/// The file size is 0 for a workbook given as bytes.
#[pyfunction]
#[pyo3(signature = (source, *, extension=None), name = "stats")]
fn compute_stats<'py>(
    py: Python<'py>,
    source: &Bound<'py, PyAny>,
    extension: Option<&str>,
) -> PyResult<Bound<'py, PyAny>> {
    let source = Source::from_py(source, extension)?;
    let computed = py.detach(|| source.stats().map_err(to_py_err))?;
    let json =
        serde_json::to_string(&computed).map_err(|e| SheetrsError::new_err(e.to_string()))?;
    py.import("json")?.call_method1("loads", (json,))
}

fn invalid(error: anyhow::Error) -> PyErr {
    PyValueError::new_err(format!("{error:#}"))
}

fn strings(values: &Bound<'_, PyTuple>) -> PyResult<Vec<String>> {
    values.extract()
}

fn scopes(values: &Bound<'_, PyTuple>) -> PyResult<Vec<ValueScope>> {
    parse_all(values)
}

/// Each argument parsed as `T` (`ValueScope`, `RowRange`, `ColumnRange`): a malformed one, or one
/// of the wrong kind (rows deleted with a column range), is a `ValueError` where it is written.
fn parse_all<T>(values: &Bound<'_, PyTuple>) -> PyResult<Vec<T>>
where
    T: std::str::FromStr<Err = anyhow::Error>,
{
    strings(values)?
        .iter()
        .map(|spec| {
            spec.parse::<T>()
                .map_err(|e| PyValueError::new_err(format!("invalid scope '{spec}': {e:#}")))
        })
        .collect()
}

/// The extensions of the workbooks the suite reads and writes.
const EXTENSIONS: [&str; 3] = ["xlsx", "xlsm", "ods"];

fn extension_of(path: &std::path::Path) -> PyResult<String> {
    path.extension()
        .and_then(|e| e.to_str())
        .map(str::to_lowercase)
        .filter(|e| EXTENSIONS.contains(&e.as_str()))
        .ok_or_else(|| {
            PyValueError::new_err(format!(
                "'{}' is not an XLSX, XLSM or ODS workbook (by its extension)",
                path.display()
            ))
        })
}

/// A workbook to edit: the file and the operations recorded on it. Nothing is read until `save`.
///
/// Every method records an operation and returns the `Workbook` that results (the one it was
/// called on is not changed), so calls chain and a workbook can be branched. Each operation
/// addresses the workbook **as the previous ones leave it**: after `delete_rows("T!3")`, `"T!3"` is
/// what was row 4. The scopes of one call are simultaneous: `delete_rows("T!3", "T!4")` deletes
/// two rows. `save(path)` runs the operations, planned into the fewest passes over the file that
/// give what running them one at a time gives, each pass reading and writing the file once as a
/// stream, so memory does not grow with the workbook. `explain()` shows the passes.
///
/// A scope that is malformed, or of the wrong kind (rows deleted with a column range), raises
/// `ValueError` where it is written. An error of the workbook (a sheet that does not exist, a 3D
/// reference, a file that cannot be read) shows when it is saved. Scopes are `"Sheet"`,
/// `"Sheet!B:D"` (columns) and `"Sheet!3:10"` (rows); quote a sheet name that contains `!`.
#[pyclass(name = "Workbook", frozen, module = "sheetrs")]
struct PyWorkbook {
    path: PathBuf,
    /// The folder that holds `path` when the workbook was given as bytes; it goes with the last
    /// handle.
    keep: Option<Arc<tempfile::TempDir>>,
    plan: EditPlan,
}

impl PyWorkbook {
    fn then(&self, operation: Operation) -> Self {
        let mut plan = self.plan.clone();
        plan.push(operation);
        Self {
            path: self.path.clone(),
            keep: self.keep.clone(),
            plan,
        }
    }
}

#[pymethods]
impl PyWorkbook {
    /// Remove sheets by name.
    #[pyo3(signature = (*names))]
    fn remove_sheets(&self, names: &Bound<'_, PyTuple>) -> PyResult<Self> {
        Ok(self.then(Operation::RemoveSheets(strings(names)?)))
    }

    /// Remove defined names (and database ranges).
    #[pyo3(signature = (*names))]
    fn remove_ranges(&self, names: &Bound<'_, PyTuple>) -> PyResult<Self> {
        Ok(self.then(Operation::RemoveNamedRanges(strings(names)?)))
    }

    /// Replace formulas by their cached values in the scopes.
    #[pyo3(signature = (*scopes))]
    fn convert_to_values(&self, scopes: &Bound<'_, PyTuple>) -> PyResult<Self> {
        Ok(self.then(Operation::ConvertToValues(self::scopes(scopes)?)))
    }

    /// Delete rows (`"Sheet!3:10"`, or `"Sheet!3"`); references to them become `#REF!`.
    #[pyo3(signature = (*rows))]
    fn delete_rows(&self, rows: &Bound<'_, PyTuple>) -> PyResult<Self> {
        Ok(self.then(Operation::DeleteRows(parse_all::<RowRange>(rows)?)))
    }

    /// Delete columns (`"Sheet!B:D"`, or `"Sheet!B"`).
    #[pyo3(signature = (*columns))]
    fn delete_columns(&self, columns: &Bound<'_, PyTuple>) -> PyResult<Self> {
        Ok(self.then(Operation::DeleteColumns(parse_all::<ColumnRange>(columns)?)))
    }

    /// The operations as they were written, as text.
    #[getter]
    fn operations(&self) -> Vec<String> {
        self.plan
            .operations()
            .iter()
            .map(ToString::to_string)
            .collect()
    }

    /// The passes over the file that run the operations, as text. Raises `ValueError` for an
    /// operation on a sheet that an earlier one removed.
    fn explain(&self) -> PyResult<String> {
        self.plan.explain().map_err(invalid)
    }

    /// Run the operations and write the workbook to `path`, which must not be the file it was
    /// opened from. With `recalc_on_open`, the application that opens the result recalculates
    /// every formula: a property of the saved file, applied once, not an operation (a conversion
    /// to values reads cached results that recalculating would drop). No file is left behind when
    /// it fails. Raises `OSError` when a file cannot be read or written and `SheetrsError` when
    /// the workbook cannot be edited or there is nothing to do.
    #[pyo3(signature = (path, *, recalc_on_open=false))]
    fn save(&self, py: Python<'_>, path: PathBuf, recalc_on_open: bool) -> PyResult<()> {
        let mut plan = self.plan.clone();
        if recalc_on_open {
            plan.recalculate_on_open();
        }
        py.detach(|| plan.execute(&self.path, &path).map_err(to_py_err))
    }

    fn __repr__(&self) -> String {
        format!(
            "Workbook({:?}, {} operation(s))",
            self.path.display().to_string(),
            self.plan.operations().len()
        )
    }
}

/// Open a workbook to edit. `source` is a path or the bytes of a workbook (then `extension` is
/// required). Only the path and the extension are checked; nothing is read until `save`, and the
/// file is never modified. Raises `FileNotFoundError` for a path that is not a file and
/// `ValueError` for one that is not an XLSX, XLSM or ODS workbook.
#[pyfunction]
#[pyo3(name = "open", signature = (source, *, extension=None))]
fn open_workbook(source: &Bound<'_, PyAny>, extension: Option<&str>) -> PyResult<PyWorkbook> {
    match Source::from_py(source, extension)? {
        Source::Path(path) => {
            extension_of(&path)?;
            if !path.is_file() {
                return Err(std::io::Error::from(std::io::ErrorKind::NotFound).into());
            }
            Ok(PyWorkbook {
                path,
                keep: None,
                plan: EditPlan::new(),
            })
        }
        Source::Bytes(bytes, extension) => {
            let extension = extension.to_lowercase();
            if !EXTENSIONS.contains(&extension.as_str()) {
                return Err(PyValueError::new_err(
                    "`extension` is 'xlsx', 'xlsm' or 'ods'",
                ));
            }
            let folder = tempfile::tempdir().map(Arc::new).map_err(|e| {
                SheetrsError::new_err(format!("cannot create a working folder: {e}"))
            })?;
            let path = folder.path().join(format!("workbook.{extension}"));
            std::fs::write(&path, bytes)?;
            Ok(PyWorkbook {
                path,
                keep: Some(folder),
                plan: EditPlan::new(),
            })
        }
    }
}

#[pymodule(gil_used = false)]
mod _sheetrs {
    use pyo3::prelude::*;

    #[pymodule_export]
    use super::{PyLinter, PyViolation, PyWorkbook, SheetrsError, compute_stats, open_workbook};

    #[pymodule_init]
    fn init(module: &Bound<'_, PyModule>) -> PyResult<()> {
        module.add("__version__", env!("CARGO_PKG_VERSION"))
    }
}

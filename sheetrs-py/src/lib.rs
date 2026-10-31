//! Python bindings for the SheetRS suite.
//!
//! The extension module is `sheetrs._sheetrs`; the `sheetrs` package re-exports it, flat. A
//! workbook is a path (`str`/`os.PathLike`), its bytes, or a `Workbook` already loaded. The
//! functions release the GIL while they read and analyze a workbook.
//!
//! - `Linter`: a configuration validated when it is built, then `lint` and `rules` (the rules
//!   exist only as part of the linter, so their catalog is next to the configuration it reflects).
//! - `stats`: the statistics of a workbook.
//! - Editing follows the eager/lazy convention of dataframe libraries: the two modes have the
//!   same methods and the same meaning (each operation addresses the workbook as the previous ones
//!   left it); they differ in what is loaded and when it runs. `read` gives a `Workbook`, with
//!   its own working copy, where each operation runs when it is called; `scan` gives a
//!   `LazyWorkbook`, where nothing is read, operations are recorded, and `sink` (to a file) or
//!   `collect` (to a `Workbook`) runs the plan, in as few passes over the file as it can.

use pyo3::exceptions::{PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList, PyTuple};
use sheetrs_core::reader::{self, Workbook};
use sheetrs_edit::EditPlan;
use sheetrs_edit::Operation;
use sheetrs_edit::writer::ValueScope;
use sheetrs_linter::report::{ViolationRecord, violation_records};
use sheetrs_linter::rules::registry;
use sheetrs_linter::{Linter, LinterConfig};
use sheetrs_stats::WorkbookStats;
use std::io::Cursor;
use std::path::{Path, PathBuf};
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
        if let Ok(workbook) = source.extract::<PyRef<'_, PyWorkbook>>() {
            return Ok(Self::Path(workbook.path.clone()));
        }
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
    strings(values)?
        .iter()
        .map(|spec| {
            spec.parse::<ValueScope>()
                .map_err(|e| PyValueError::new_err(format!("invalid scope '{spec}': {e:#}")))
        })
        .collect()
}

/// The extensions of the workbooks the suite reads and writes.
const EXTENSIONS: [&str; 3] = ["xlsx", "xlsm", "ods"];

fn extension_of(path: &Path) -> PyResult<String> {
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

fn temporary_folder() -> PyResult<Arc<tempfile::TempDir>> {
    tempfile::tempdir()
        .map(Arc::new)
        .map_err(|e| SheetrsError::new_err(format!("cannot create a working folder: {e}")))
}

/// The operations of the two modes, built from the arguments of a call.
fn remove_sheets_op(names: &Bound<'_, PyTuple>) -> PyResult<Operation> {
    Ok(Operation::RemoveSheets(strings(names)?))
}
fn remove_ranges_op(names: &Bound<'_, PyTuple>) -> PyResult<Operation> {
    Ok(Operation::RemoveNamedRanges(strings(names)?))
}
fn convert_op(specs: &Bound<'_, PyTuple>) -> PyResult<Operation> {
    Ok(Operation::ConvertToValues(scopes(specs)?))
}
fn delete_rows_op(specs: &Bound<'_, PyTuple>) -> PyResult<Operation> {
    Ok(Operation::DeleteRows(scopes(specs)?))
}
fn delete_columns_op(specs: &Bound<'_, PyTuple>) -> PyResult<Operation> {
    Ok(Operation::DeleteColumns(scopes(specs)?))
}

/// A workbook loaded for editing: eager mode.
///
/// It has a working copy of the file, and every operation runs when it is called, on the result
/// of the previous ones, and returns the `Workbook` that results (the one it was called on is not
/// changed). So an error shows where it is made, the intermediate workbook can be linted or
/// measured (`Linter.lint(workbook)`, `stats(workbook)`), and a position is the one the earlier
/// operations left. Each operation reads and writes the file once; `lazy()` and the plan of a
/// `LazyWorkbook` do the same work in fewer passes. `write` saves it.
///
/// Scopes are `"Sheet"`, `"Sheet!B:D"` (columns) and `"Sheet!3:10"` (rows); quote a sheet name
/// that contains `!`. The scopes of one call are simultaneous.
#[pyclass(name = "Workbook", frozen, module = "sheetrs")]
struct PyWorkbook {
    path: PathBuf,
    /// The folder that holds `path`, when it is a working copy; it goes with the last handle.
    keep: Option<Arc<tempfile::TempDir>>,
}

impl PyWorkbook {
    fn apply(&self, py: Python<'_>, operation: Operation) -> PyResult<Self> {
        let mut plan = EditPlan::new();
        plan.push(operation).map_err(invalid)?;
        if plan.is_empty() {
            return Ok(Self {
                path: self.path.clone(),
                keep: self.keep.clone(),
            });
        }
        let folder = temporary_folder()?;
        let output = folder
            .path()
            .join(format!("workbook.{}", extension_of(&self.path)?));
        py.detach(|| plan.execute(&self.path, &output).map_err(to_py_err))?;
        Ok(Self {
            path: output,
            keep: Some(folder),
        })
    }
}

#[pymethods]
impl PyWorkbook {
    /// Remove sheets by name.
    #[pyo3(signature = (*names))]
    fn remove_sheets(&self, py: Python<'_>, names: &Bound<'_, PyTuple>) -> PyResult<Self> {
        self.apply(py, remove_sheets_op(names)?)
    }

    /// Remove defined names (and database ranges).
    #[pyo3(signature = (*names))]
    fn remove_ranges(&self, py: Python<'_>, names: &Bound<'_, PyTuple>) -> PyResult<Self> {
        self.apply(py, remove_ranges_op(names)?)
    }

    /// Replace formulas by their cached values in the scopes.
    #[pyo3(signature = (*scopes))]
    fn convert_to_values(&self, py: Python<'_>, scopes: &Bound<'_, PyTuple>) -> PyResult<Self> {
        self.apply(py, convert_op(scopes)?)
    }

    /// Delete rows (`"Sheet!3:10"`, or `"Sheet!3"`); references to them become `#REF!`.
    #[pyo3(signature = (*scopes))]
    fn delete_rows(&self, py: Python<'_>, scopes: &Bound<'_, PyTuple>) -> PyResult<Self> {
        self.apply(py, delete_rows_op(scopes)?)
    }

    /// Delete columns (`"Sheet!B:D"`, or `"Sheet!B"`).
    #[pyo3(signature = (*scopes))]
    fn delete_columns(&self, py: Python<'_>, scopes: &Bound<'_, PyTuple>) -> PyResult<Self> {
        self.apply(py, delete_columns_op(scopes)?)
    }

    /// Mark the workbook so the application that opens it recalculates every formula.
    fn recalc_on_open(&self, py: Python<'_>) -> PyResult<Self> {
        self.apply(py, Operation::RecalculateOnOpen)
    }

    /// The same workbook as a `LazyWorkbook`, with no operations: what follows is recorded, not run.
    fn lazy(&self) -> PyLazyWorkbook {
        PyLazyWorkbook {
            path: self.path.clone(),
            keep: self.keep.clone(),
            plan: EditPlan::new(),
        }
    }

    /// Save the workbook to `path`, which must not be the file it was read from. Raises `OSError`
    /// when it cannot be written.
    fn write(&self, py: Python<'_>, path: PathBuf) -> PyResult<()> {
        let same = std::fs::canonicalize(&path)
            .ok()
            .zip(std::fs::canonicalize(&self.path).ok())
            .is_some_and(|(a, b)| a == b);
        if same {
            return Err(SheetrsError::new_err(format!(
                "'{}' is the file this workbook is read from",
                path.display()
            )));
        }
        py.detach(|| std::fs::copy(&self.path, &path).map(|_| ()))?;
        Ok(())
    }

    fn __repr__(&self) -> String {
        format!("Workbook({:?})", self.path.display().to_string())
    }
}

/// A workbook scanned for editing: lazy mode. Nothing is read.
///
/// It has the methods of `Workbook` with the same meaning, each operation addressing the workbook
/// as the previous ones leave it, but they only record the operation and return the plan that
/// results (the one they were called on is not changed). The plan runs on `sink(path)`, which
/// writes the file, or `collect()`, which gives a `Workbook`. It is planned then: the operations
/// are grouped into the fewest passes that give what running them one at a time gives, each pass
/// reading and writing the file once as a stream, so memory does not grow with the workbook.
/// `explain()` shows the passes. An error of the workbook shows when the plan runs; one of the
/// plan itself (an operation on a sheet an earlier one removed) when it is planned.
#[pyclass(name = "LazyWorkbook", frozen, module = "sheetrs")]
struct PyLazyWorkbook {
    path: PathBuf,
    keep: Option<Arc<tempfile::TempDir>>,
    plan: EditPlan,
}

impl PyLazyWorkbook {
    fn then(&self, operation: Operation) -> PyResult<Self> {
        let mut plan = self.plan.clone();
        plan.push(operation).map_err(invalid)?;
        Ok(Self {
            path: self.path.clone(),
            keep: self.keep.clone(),
            plan,
        })
    }
}

#[pymethods]
impl PyLazyWorkbook {
    /// Remove sheets by name.
    #[pyo3(signature = (*names))]
    fn remove_sheets(&self, names: &Bound<'_, PyTuple>) -> PyResult<Self> {
        self.then(remove_sheets_op(names)?)
    }

    /// Remove defined names (and database ranges).
    #[pyo3(signature = (*names))]
    fn remove_ranges(&self, names: &Bound<'_, PyTuple>) -> PyResult<Self> {
        self.then(remove_ranges_op(names)?)
    }

    /// Replace formulas by their cached values in the scopes.
    #[pyo3(signature = (*scopes))]
    fn convert_to_values(&self, scopes: &Bound<'_, PyTuple>) -> PyResult<Self> {
        self.then(convert_op(scopes)?)
    }

    /// Delete rows (`"Sheet!3:10"`, or `"Sheet!3"`); references to them become `#REF!`.
    #[pyo3(signature = (*scopes))]
    fn delete_rows(&self, scopes: &Bound<'_, PyTuple>) -> PyResult<Self> {
        self.then(delete_rows_op(scopes)?)
    }

    /// Delete columns (`"Sheet!B:D"`, or `"Sheet!B"`).
    #[pyo3(signature = (*scopes))]
    fn delete_columns(&self, scopes: &Bound<'_, PyTuple>) -> PyResult<Self> {
        self.then(delete_columns_op(scopes)?)
    }

    /// Mark the workbook so the application that opens it recalculates every formula.
    fn recalc_on_open(&self) -> PyResult<Self> {
        self.then(Operation::RecalculateOnOpen)
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

    /// The passes over the file that run the plan, as text.
    fn explain(&self) -> PyResult<String> {
        self.plan.explain().map_err(invalid)
    }

    /// Run the plan and write the result to `path`, which must not be the file it reads. No file
    /// is left behind when it fails. Raises `OSError` when a file cannot be read or written and
    /// `SheetrsError` when the workbook cannot be edited or the plan has no operations.
    fn sink(&self, py: Python<'_>, path: PathBuf) -> PyResult<()> {
        py.detach(|| self.plan.execute(&self.path, &path).map_err(to_py_err))
    }

    /// Run the plan and load the result: a `Workbook`.
    fn collect(&self, py: Python<'_>) -> PyResult<PyWorkbook> {
        let folder = temporary_folder()?;
        let output = folder
            .path()
            .join(format!("workbook.{}", extension_of(&self.path)?));
        py.detach(|| self.plan.execute(&self.path, &output).map_err(to_py_err))?;
        Ok(PyWorkbook {
            path: output,
            keep: Some(folder),
        })
    }

    fn __repr__(&self) -> String {
        format!(
            "LazyWorkbook({:?}, {} operation(s))",
            self.path.display().to_string(),
            self.plan.operations().len()
        )
    }
}

/// Load a workbook for editing, eagerly: a `Workbook` whose operations run when called.
///
/// `source` is a path or the bytes of a workbook (then `extension` is required). The file is
/// checked (it exists, it is a zip archive of a workbook); it is never modified, whatever is
/// done with the `Workbook`. Raises `OSError` when it cannot be read.
#[pyfunction]
#[pyo3(signature = (source, *, extension=None))]
fn read(
    py: Python<'_>,
    source: &Bound<'_, PyAny>,
    extension: Option<&str>,
) -> PyResult<PyWorkbook> {
    match Source::from_py(source, extension)? {
        Source::Path(path) => {
            extension_of(&path)?;
            py.detach(|| {
                let file = std::fs::File::open(&path)?;
                zip::ZipArchive::new(file)
                    .map(|_| ())
                    .map_err(|e| SheetrsError::new_err(format!("{}: {e}", path.display())))
                    .map(|()| PyWorkbook { path, keep: None })
            })
        }
        Source::Bytes(bytes, extension) => {
            let extension = extension.to_lowercase();
            if !EXTENSIONS.contains(&extension.as_str()) {
                return Err(PyValueError::new_err(
                    "`extension` is 'xlsx', 'xlsm' or 'ods'",
                ));
            }
            zip::ZipArchive::new(Cursor::new(&bytes))
                .map_err(|e| SheetrsError::new_err(format!("not a workbook: {e}")))?;
            let folder = temporary_folder()?;
            let path = folder.path().join(format!("workbook.{extension}"));
            std::fs::write(&path, bytes)?;
            Ok(PyWorkbook {
                path,
                keep: Some(folder),
            })
        }
    }
}

/// Scan a workbook for editing, lazily: a `LazyWorkbook` that reads nothing until its plan runs.
///
/// `path` must exist and be an XLSX, XLSM or ODS file; that is all that is checked.
#[pyfunction]
fn scan(path: PathBuf) -> PyResult<PyLazyWorkbook> {
    extension_of(&path)?;
    if !path.is_file() {
        return Err(std::io::Error::from(std::io::ErrorKind::NotFound).into());
    }
    Ok(PyLazyWorkbook {
        path,
        keep: None,
        plan: EditPlan::new(),
    })
}

#[pymodule(gil_used = false)]
mod _sheetrs {
    use pyo3::prelude::*;

    #[pymodule_export]
    use super::{
        PyLazyWorkbook, PyLinter, PyViolation, PyWorkbook, SheetrsError, compute_stats, read, scan,
    };

    #[pymodule_init]
    fn init(module: &Bound<'_, PyModule>) -> PyResult<()> {
        module.add("__version__", env!("CARGO_PKG_VERSION"))
    }
}

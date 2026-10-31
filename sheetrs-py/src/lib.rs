//! Python bindings for the SheetRS suite.
//!
//! The extension module is `sheetrs._sheetrs`; the `sheetrs` package re-exports it, flat: the
//! three tools are `Linter`/`lint`, `stats` and `EditPlan`/`edit`. A workbook is a path
//! (`str`/`os.PathLike`) or its bytes. The functions release the GIL while they read and analyze
//! a workbook.
//!
//! - `Linter`: a configuration validated when it is built, then `lint` and `rules` (the rules
//!   exist only as part of the linter, so their catalog is next to the configuration it reflects).
//! - `EditPlan`: lazy editing. Operations are collected, each written against the input workbook,
//!   and `save` applies them resolved into one order in a single pass. `edit` is the eager form:
//!   the same operations given at once, applied at once.

use pyo3::exceptions::{PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList, PyTuple};
use sheetrs_core::reader::{self, Workbook};
use sheetrs_edit::EditPlan;
use sheetrs_edit::writer::ValueScope;
use sheetrs_linter::report::{ViolationRecord, violation_records};
use sheetrs_linter::rules::registry;
use sheetrs_linter::{Linter, LinterConfig};
use sheetrs_stats::WorkbookStats;
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

/// Lint a workbook with a one-off linter: `Linter(config=…, config_path=…).lint(source, …)`.
#[pyfunction]
#[pyo3(signature = (source, *, config=None, config_path=None, extension=None))]
fn lint(
    py: Python<'_>,
    source: &Bound<'_, PyAny>,
    config: Option<&str>,
    config_path: Option<PathBuf>,
    extension: Option<&str>,
) -> PyResult<Vec<PyViolation>> {
    PyLinter::new(config, config_path)?.lint(py, source, extension)
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

/// Operations on one workbook, written against the input workbook and applied in a single pass.
///
/// Lazy mode: each method records operations and returns the plan, so calls chain; nothing is
/// read or written until `save(output)`. Operations address the rows, columns and sheet names of
/// the **input**, in any order: the plan merges overlapping row and column ranges and orders what
/// it does so that no operation sees the positions another one moved (formulas are converted
/// before the cells they read are deleted, rows and columns are deleted from the last to the
/// first, sheets are removed last). `operations` shows that order. The result does not depend on
/// the order the operations were added in, and equals running them one per pass. Adding an
/// operation twice counts once. A malformed scope, or rows deleted with a column range, raises
/// `ValueError` where it is added. A formula with a 3D reference (`Sheet1:Sheet3!A1`) makes
/// `save` fail when rows, columns or sheets are deleted or removed: not supported yet (issue #75).
///
#[pyclass(name = "EditPlan", module = "sheetrs")]
struct PyEditPlan {
    input: PathBuf,
    plan: EditPlan,
}

#[pymethods]
impl PyEditPlan {
    #[new]
    fn new(input: PathBuf) -> Self {
        Self {
            input,
            plan: EditPlan::new(),
        }
    }

    /// Remove sheets by name.
    #[pyo3(signature = (*names))]
    fn remove_sheets<'py>(
        mut slf: PyRefMut<'py, Self>,
        names: &Bound<'py, PyTuple>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        for name in strings(names)? {
            slf.plan.remove_sheet(name);
        }
        Ok(slf)
    }

    /// Remove defined names (and database ranges).
    #[pyo3(signature = (*names))]
    fn remove_ranges<'py>(
        mut slf: PyRefMut<'py, Self>,
        names: &Bound<'py, PyTuple>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        for name in strings(names)? {
            slf.plan.remove_named_range(name);
        }
        Ok(slf)
    }

    /// Replace formulas by their cached values in scopes: `"Sheet"`, `"Sheet!B:D"`, `"Sheet!3:10"`.
    #[pyo3(signature = (*scopes))]
    fn convert_to_values<'py>(
        mut slf: PyRefMut<'py, Self>,
        scopes: &Bound<'py, PyTuple>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        for scope in self::scopes(scopes)? {
            slf.plan.convert_to_values(scope);
        }
        Ok(slf)
    }

    /// Delete rows (`"Sheet!3:10"`, or `"Sheet!3"`); references to them become `#REF!`.
    #[pyo3(signature = (*scopes))]
    fn delete_rows<'py>(
        mut slf: PyRefMut<'py, Self>,
        scopes: &Bound<'py, PyTuple>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        for scope in self::scopes(scopes)? {
            slf.plan.delete_rows(scope).map_err(invalid)?;
        }
        Ok(slf)
    }

    /// Delete columns (`"Sheet!B:D"`, or `"Sheet!B"`).
    #[pyo3(signature = (*scopes))]
    fn delete_columns<'py>(
        mut slf: PyRefMut<'py, Self>,
        scopes: &Bound<'py, PyTuple>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        for scope in self::scopes(scopes)? {
            slf.plan.delete_columns(scope).map_err(invalid)?;
        }
        Ok(slf)
    }

    /// Mark the output so the application that opens it recalculates every formula.
    fn recalc_on_open(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.plan.recalculate_on_open();
        slf
    }

    /// The operations as they will be applied, as text: resolved (overlapping ranges merged,
    /// repeats and scopes on removed sheets dropped) and in the order that gives the right result,
    /// whatever the order they were added in.
    #[getter]
    fn operations(&self) -> Vec<String> {
        self.plan
            .operations()
            .iter()
            .map(ToString::to_string)
            .collect()
    }

    /// Apply the operations to the input in one pass and write `output`, which must not be the
    /// input. The plan is not consumed: it can be saved again. Raises `OSError` when a file
    /// cannot be read or written and `SheetrsError` when the operation fails (no output is left
    /// behind) or the plan is empty.
    fn save(&self, py: Python<'_>, output: PathBuf) -> PyResult<()> {
        py.detach(|| self.plan.execute(&self.input, &output).map_err(to_py_err))
    }

    fn __len__(&self) -> usize {
        self.plan.operations().len()
    }

    fn __repr__(&self) -> String {
        format!(
            "Plan({:?}, {} operation(s))",
            self.input.display(),
            self.plan.operations().len()
        )
    }
}

/// Eager editing: write `output` from `input` with the operations given, now. It is
/// `EditPlan(input)` with each argument added and `save(output)`; every argument is a list.
///
/// - `remove_sheets`, `remove_ranges`: names to remove.
/// - `convert_to_values`: scopes (`"Sheet"`, `"Sheet!B:D"`, `"Sheet!3:10"`) whose formulas are
///   replaced by their cached values.
/// - `delete_rows` (`"Sheet!3:10"`), `delete_columns` (`"Sheet!B:D"`): of the input; references to
///   them become `#REF!`.
/// - `recalc_on_open`: the application that opens the output recalculates every formula.
///
/// Raises `ValueError` for a malformed scope and `SheetrsError` when the operation fails (no
/// output is left behind) or nothing was asked.
#[pyfunction]
#[pyo3(signature = (
    input,
    output,
    *,
    remove_sheets=None,
    remove_ranges=None,
    convert_to_values=None,
    delete_rows=None,
    delete_columns=None,
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
    delete_columns: Option<Vec<String>>,
    recalc_on_open: bool,
) -> PyResult<()> {
    fn parse(specs: Option<Vec<String>>) -> PyResult<Vec<ValueScope>> {
        specs
            .unwrap_or_default()
            .iter()
            .map(|spec| {
                spec.parse::<ValueScope>()
                    .map_err(|e| PyValueError::new_err(format!("invalid scope '{spec}': {e:#}")))
            })
            .collect()
    }
    let mut plan = EditPlan::new();
    for name in remove_sheets.unwrap_or_default() {
        plan.remove_sheet(name);
    }
    for name in remove_ranges.unwrap_or_default() {
        plan.remove_named_range(name);
    }
    for scope in parse(convert_to_values)? {
        plan.convert_to_values(scope);
    }
    for scope in parse(delete_rows)? {
        plan.delete_rows(scope).map_err(invalid)?;
    }
    for scope in parse(delete_columns)? {
        plan.delete_columns(scope).map_err(invalid)?;
    }
    if recalc_on_open {
        plan.recalculate_on_open();
    }
    py.detach(|| plan.execute(&input, &output).map_err(to_py_err))
}

#[pymodule(gil_used = false)]
mod _sheetrs {
    use pyo3::prelude::*;

    #[pymodule_export]
    use super::{PyEditPlan, PyLinter, PyViolation, SheetrsError, compute_stats, edit, lint};

    #[pymodule_init]
    fn init(module: &Bound<'_, PyModule>) -> PyResult<()> {
        module.add("__version__", env!("CARGO_PKG_VERSION"))
    }
}

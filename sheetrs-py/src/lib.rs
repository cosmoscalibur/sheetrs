//! Python bindings for the SheetRS suite.
//!
//! The extension module is `sheetrs._sheetrs`; the `sheetrs` package re-exports its submodules.
//! There is one submodule per tool, as there is one crate per tool: the rules exist only as part
//! of the linter, so their catalog is `Linter.rules()`, next to the configuration it reflects.
//!
//! - `linter` (`sheetrs-linter`): `Linter` (validated configuration, `lint`, `rules`), `lint`,
//!   `Violation`.
//! - `stats` (`sheetrs-stats`): `compute`.
//! - `edit` (`sheetrs-edit`): `Plan` (lazy: operations collected, applied in one pass on `save`)
//!   and one function per operation (eager: applied at once).
//!
//! A workbook is a path (`str`/`os.PathLike`) or its bytes. The functions release the GIL while
//! they read and analyze a workbook.

use pyo3::exceptions::{PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList, PyTuple};
use sheetrs_core::reader::{self, Workbook};
use sheetrs_edit::EditPlan;
use sheetrs_edit::writer::ValueScope;
use sheetrs_linter::report::{ViolationRecord, violation_records};
use sheetrs_linter::rules::registry;
use sheetrs_linter::{Linter, LinterConfig};
use sheetrs_stats::stats::WorkbookStats;
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
#[pyclass(name = "Linter", frozen, module = "sheetrs.linter")]
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
#[pyo3(signature = (source, *, extension=None))]
#[pyo3(name = "compute")]
fn stats_compute<'py>(
    py: Python<'py>,
    source: &Bound<'py, PyAny>,
    extension: Option<&str>,
) -> PyResult<Bound<'py, PyAny>> {
    let source = Source::from_py(source, extension)?;
    let stats = py.detach(|| source.stats().map_err(to_py_err))?;
    let json = serde_json::to_string(&stats).map_err(|e| SheetrsError::new_err(e.to_string()))?;
    py.import("json")?.call_method1("loads", (json,))
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

/// Operations on one workbook, collected and applied together.
///
/// Lazy mode: each method records operations and returns the plan, so calls chain; nothing is
/// read or written until `save(output)`, which applies all of them while it streams the file
/// once. All the operations address the **input** workbook (its rows, columns and sheet names),
/// whatever order they were added in; adding an operation twice counts once. A malformed scope
/// raises `ValueError` where it is added.
///
/// For immediate execution, one pass per operation, see the functions of this module.
#[pyclass(name = "Plan", module = "sheetrs.edit")]
struct PyPlan {
    input: PathBuf,
    plan: EditPlan,
}

#[pymethods]
impl PyPlan {
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
            slf.plan.delete_rows(scope);
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
            slf.plan.delete_columns(scope);
        }
        Ok(slf)
    }

    /// Mark the output so the application that opens it recalculates every formula.
    fn recalc_on_open(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.plan.recalculate_on_open();
        slf
    }

    /// The recorded operations, in the order they were added, as text.
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

/// Eager mode: runs one operation on `input` now and writes `output`.
fn run_eager(
    py: Python<'_>,
    input: PathBuf,
    output: PathBuf,
    add: impl FnOnce(&mut EditPlan) -> PyResult<()>,
) -> PyResult<()> {
    let mut plan = EditPlan::new();
    add(&mut plan)?;
    py.detach(|| plan.execute(&input, &output).map_err(to_py_err))
}

/// Eager: write `output` without the sheets `names`.
#[pyfunction]
#[pyo3(signature = (input, output, *names), name = "remove_sheets")]
fn edit_remove_sheets(
    py: Python<'_>,
    input: PathBuf,
    output: PathBuf,
    names: &Bound<'_, PyTuple>,
) -> PyResult<()> {
    run_eager(py, input, output, |plan| {
        strings(names)?.into_iter().for_each(|n| {
            plan.remove_sheet(n);
        });
        Ok(())
    })
}

/// Eager: write `output` without the defined names `names`.
#[pyfunction]
#[pyo3(signature = (input, output, *names), name = "remove_ranges")]
fn edit_remove_ranges(
    py: Python<'_>,
    input: PathBuf,
    output: PathBuf,
    names: &Bound<'_, PyTuple>,
) -> PyResult<()> {
    run_eager(py, input, output, |plan| {
        strings(names)?.into_iter().for_each(|n| {
            plan.remove_named_range(n);
        });
        Ok(())
    })
}

/// Eager: write `output` with the formulas of `scopes` replaced by their cached values.
#[pyfunction]
#[pyo3(signature = (input, output, *scopes), name = "convert_to_values")]
fn edit_convert_to_values(
    py: Python<'_>,
    input: PathBuf,
    output: PathBuf,
    scopes: &Bound<'_, PyTuple>,
) -> PyResult<()> {
    run_eager(py, input, output, |plan| {
        self::scopes(scopes)?.into_iter().for_each(|s| {
            plan.convert_to_values(s);
        });
        Ok(())
    })
}

/// Eager: write `output` without the rows `scopes` (`"Sheet!3:10"`).
#[pyfunction]
#[pyo3(signature = (input, output, *scopes), name = "delete_rows")]
fn edit_delete_rows(
    py: Python<'_>,
    input: PathBuf,
    output: PathBuf,
    scopes: &Bound<'_, PyTuple>,
) -> PyResult<()> {
    run_eager(py, input, output, |plan| {
        self::scopes(scopes)?.into_iter().for_each(|s| {
            plan.delete_rows(s);
        });
        Ok(())
    })
}

/// Eager: write `output` without the columns `scopes` (`"Sheet!B:D"`).
#[pyfunction]
#[pyo3(signature = (input, output, *scopes), name = "delete_columns")]
fn edit_delete_columns(
    py: Python<'_>,
    input: PathBuf,
    output: PathBuf,
    scopes: &Bound<'_, PyTuple>,
) -> PyResult<()> {
    run_eager(py, input, output, |plan| {
        self::scopes(scopes)?.into_iter().for_each(|s| {
            plan.delete_columns(s);
        });
        Ok(())
    })
}

/// Eager: write `output` marked to be recalculated when opened.
#[pyfunction]
#[pyo3(name = "recalc_on_open")]
fn edit_recalc_on_open(py: Python<'_>, input: PathBuf, output: PathBuf) -> PyResult<()> {
    run_eager(py, input, output, |plan| {
        plan.recalculate_on_open();
        Ok(())
    })
}

#[pymodule(gil_used = false)]
mod _sheetrs {
    use pyo3::prelude::*;

    #[pymodule_export]
    use super::SheetrsError;

    #[pymodule]
    mod linter {
        #[pymodule_export]
        use super::super::{PyLinter, PyViolation, lint};
    }

    #[pymodule]
    mod stats {
        #[pymodule_export]
        use super::super::stats_compute;
    }

    #[pymodule]
    mod edit {
        #[pymodule_export]
        use super::super::{
            PyPlan, edit_convert_to_values, edit_delete_columns, edit_delete_rows,
            edit_recalc_on_open, edit_remove_ranges, edit_remove_sheets,
        };
    }

    #[pymodule_init]
    fn init(module: &Bound<'_, PyModule>) -> PyResult<()> {
        module.add("__version__", env!("CARGO_PKG_VERSION"))
    }
}

//! Python bindings for the `SheetRS` suite.
//!
//! The extension module is `sheetrs._sheetrs`, flat; the `sheetrs` package presents it as the
//! crates are laid out: `sheetrs` (core: `open`, `Source`, `SheetrsError`), `sheetrs.linter`,
//! `sheetrs.stats` and `sheetrs.edit`.
//!
//! `open` gives the input of every tool: a [`Source`], a reference to a workbook that reads
//! nothing and is not a model of it, and without behavior of its own. The tool is the subject and
//! takes it, as in Rust:
//!
//! - `Linter(...).lint(source)`: an object, because it has a configuration.
//! - `stats.collect(source)`: a function, because there is nothing to configure.
//! - `Editor(source).delete_rows(...).save(path)`: an object, because it accumulates changes.
//!
//! The functions release the GIL while they read and analyze a workbook.

use pyo3::exceptions::{PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList, PyTuple};
use sheetrs_edit::{ColumnRange, Editor, RowRange, ValueScope};
use sheetrs_io::reader::{Format, Source};
use sheetrs_linter::report::{ViolationRecord, violation_records};
use sheetrs_linter::rules::registry;
use sheetrs_linter::{Linter, LinterConfig};
use std::path::PathBuf;

pyo3::create_exception!(
    _sheetrs,
    SheetrsError,
    pyo3::exceptions::PyException,
    "A workbook could not be read, analyzed or written."
);

/// A failure of the file system anywhere in the chain of causes becomes `OSError`
/// (`FileNotFoundError`, …), keeping its kind; anything else, `SheetrsError`. The message is the
/// whole chain.
fn to_py_err<E: std::error::Error + 'static>(error: E) -> PyErr {
    let mut cause: Option<&(dyn std::error::Error + 'static)> = Some(&error);
    while let Some(current) = cause {
        if let Some(io) = current.downcast_ref::<std::io::Error>() {
            return PyErr::from(std::io::Error::new(io.kind(), error.to_string()));
        }
        cause = current.source();
    }
    SheetrsError::new_err(error.to_string())
}

/// One finding of the linter.
#[pyclass(name = "Violation", frozen, module = "sheetrs.linter")]
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

    /// Lint the workbook of a `Source` (the result of `sheetrs.open`). Returns the violations
    /// ordered by scope (book, sheet, cell) and rule.
    ///
    /// Raises `OSError` when a file cannot be read and `SheetrsError` when the workbook cannot be
    /// parsed.
    fn lint(&self, py: Python<'_>, source: &PySource) -> PyResult<Vec<PyViolation>> {
        let source = &source.source;
        py.detach(|| {
            let workbook = source.read_workbook().map_err(to_py_err)?;
            let violations = self.linter.lint_workbook(&workbook);
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
}

/// The input of every tool: a reference to a workbook (a file, or its bytes), made by `open`.
///
/// It reads nothing and is not a model of the workbook: the tools read it when they run, and the
/// same `Source` can be given to several. It has no behavior of its own: `Linter.lint(source)`,
/// `stats.collect(source)` and `Editor(source)` are the tools that take it.
#[pyclass(name = "Source", frozen, module = "sheetrs")]
struct PySource {
    source: Source,
}

#[pymethods]
impl PySource {
    fn __repr__(&self) -> String {
        format!("Source({})", self.source)
    }
}

/// The statistics of the workbook of a `Source`, as `sheetstats --format json` reports them:
/// totals plus a `sheets` list with one dict per sheet. The file size is 0 for a workbook given
/// as bytes. Raises `OSError` when a file cannot be read and `SheetrsError` when the workbook
/// cannot be parsed.
#[pyfunction]
#[pyo3(name = "collect")]
fn collect_stats<'py>(py: Python<'py>, source: &PySource) -> PyResult<Bound<'py, PyAny>> {
    let source = &source.source;
    let computed = py.detach(|| sheetrs_stats::collect(source).map_err(to_py_err))?;
    let json =
        serde_json::to_string(&computed).map_err(|e| SheetrsError::new_err(e.to_string()))?;
    py.import("json")?.call_method1("loads", (json,))
}

/// The changes to make to a workbook, and the way to save them.
///
/// An editor is made for a `Source` and collects changes: every method adds one and returns the
/// editor, so calls chain. Each change addresses the workbook **as the previous ones leave it**:
/// after `delete_rows("T!3")`, `"T!3"` is what was row 4. The scopes of one call are simultaneous:
/// `delete_rows("T!3", "T!4")` deletes two rows. Nothing is read until `save(path)`, which runs
/// the changes in the fewest passes over the file that give what running them one at a time
/// gives; each pass reads and writes the file once as a stream, so memory does not grow with the
/// workbook. `explain()` says what `save` will do.
///
/// A scope that is malformed, or of the wrong kind (rows deleted with a column range), raises
/// `ValueError` where it is written. An error of the workbook (a sheet that does not exist, a 3D
/// reference, a file that cannot be read) shows when it is saved. Scopes are `"Sheet"`,
/// `"Sheet!B:D"` (columns) and `"Sheet!3:10"` (rows); quote a sheet name that contains `!`.
#[pyclass(name = "Editor", module = "sheetrs.edit")]
struct PyEditor {
    editor: Editor,
}

#[pymethods]
impl PyEditor {
    #[new]
    fn new(source: &PySource) -> Self {
        Self {
            editor: Editor::new(&source.source),
        }
    }

    /// Remove sheets by name.
    #[pyo3(signature = (*names))]
    fn remove_sheets<'py>(
        mut slf: PyRefMut<'py, Self>,
        names: &Bound<'_, PyTuple>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor.remove_sheets(strings(names)?);
        Ok(slf)
    }

    /// Remove defined names (and database ranges).
    #[pyo3(signature = (*names))]
    fn remove_ranges<'py>(
        mut slf: PyRefMut<'py, Self>,
        names: &Bound<'_, PyTuple>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor.remove_named_ranges(strings(names)?);
        Ok(slf)
    }

    /// Replace formulas by their cached values in the scopes.
    #[pyo3(signature = (*scopes))]
    fn convert_to_values<'py>(
        mut slf: PyRefMut<'py, Self>,
        scopes: &Bound<'_, PyTuple>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .convert_to_values(parse_all::<ValueScope>(scopes)?);
        Ok(slf)
    }

    /// Delete rows (`"Sheet!3:10"`, or `"Sheet!3"`); references to them become `#REF!`.
    #[pyo3(signature = (*rows))]
    fn delete_rows<'py>(
        mut slf: PyRefMut<'py, Self>,
        rows: &Bound<'_, PyTuple>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor.delete_rows(parse_all::<RowRange>(rows)?);
        Ok(slf)
    }

    /// Delete columns (`"Sheet!B:D"`, or `"Sheet!B"`).
    #[pyo3(signature = (*columns))]
    fn delete_columns<'py>(
        mut slf: PyRefMut<'py, Self>,
        columns: &Bound<'_, PyTuple>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .delete_columns(parse_all::<ColumnRange>(columns)?);
        Ok(slf)
    }

    /// What `save` will do: the passes over the file that run the changes, one change per line,
    /// and the mark to recalculate when there is one. Raises `ValueError` for a change on a sheet
    /// that an earlier one removed.
    fn explain(&self) -> PyResult<String> {
        self.editor.explain().map_err(invalid)
    }

    /// Run the changes and write the workbook to `path`, which must not be the file it was
    /// opened from. No file is left behind when it fails. Raises `OSError` when a file cannot be
    /// read or written and `SheetrsError` when the workbook cannot be edited (including when
    /// there are no changes).
    ///
    /// The editor decides whether the saved file is marked so that the application that opens it
    /// recalculates every formula, and puts the mark last: it does when a change leaves the
    /// cached results of formulas stale (rows, columns, sheets or names removed). There is no
    /// option for it.
    #[allow(clippy::needless_pass_by_value)] // pyo3 hands over an owned argument
    fn save(&self, py: Python<'_>, path: std::borrow::Cow<'_, std::path::Path>) -> PyResult<()> {
        py.detach(|| self.editor.save(&path).map(|_| ()).map_err(to_py_err))
    }

    fn __repr__(&self) -> String {
        let marks = if self.editor.is_empty() {
            "no changes"
        } else {
            "changes"
        };
        format!("Editor({marks})")
    }
}

fn strings(values: &Bound<'_, PyTuple>) -> PyResult<Vec<String>> {
    values.extract()
}

fn invalid(error: impl std::fmt::Display) -> PyErr {
    PyValueError::new_err(error.to_string())
}

/// Each argument parsed as `T` (`ValueScope`, `RowRange`, `ColumnRange`): a malformed one, or one
/// of the wrong kind (rows deleted with a column range), is a `ValueError` where it is written.
fn parse_all<T>(values: &Bound<'_, PyTuple>) -> PyResult<Vec<T>>
where
    T: std::str::FromStr,
    T::Err: std::fmt::Display,
{
    strings(values)?
        .iter()
        .map(|spec| {
            spec.parse::<T>()
                .map_err(|e| PyValueError::new_err(format!("invalid scope '{spec}': {e}")))
        })
        .collect()
}

/// Open a workbook: make the `Source` that every tool takes. `source` is a path or the bytes of a
/// workbook (then `extension` is required). Nothing is read, and the file is never modified.
/// Raises `FileNotFoundError` for a path that is not a file and `ValueError` for one that is not an
/// XLSX, XLSM or ODS workbook.
#[pyfunction]
#[pyo3(name = "open", signature = (source, *, extension=None))]
fn open_workbook(source: &Bound<'_, PyAny>, extension: Option<&str>) -> PyResult<PySource> {
    if let Ok(bytes) = source.cast::<PyBytes>() {
        let Some(extension) = extension else {
            return Err(PyValueError::new_err(
                "`extension` ('xlsx', 'xlsm' or 'ods') is required to open a workbook from bytes",
            ));
        };
        let format: Format = extension.parse().map_err(invalid)?;
        return Ok(PySource {
            source: Source::from_bytes(bytes.as_bytes().to_vec(), format),
        });
    }
    let path = source
        .extract::<PathBuf>()
        .map_err(|_| PyTypeError::new_err("a workbook is a path (str or os.PathLike) or bytes"))?;
    let source = Source::from_file(path).map_err(invalid)?;
    if source.path().is_some_and(|path| !path.is_file()) {
        return Err(std::io::Error::from(std::io::ErrorKind::NotFound).into());
    }
    Ok(PySource { source })
}

#[pymodule(gil_used = false)]
mod _sheetrs {
    use pyo3::prelude::*;

    #[pymodule_export]
    use super::{
        PyEditor, PyLinter, PySource, PyViolation, SheetrsError, collect_stats, open_workbook,
    };

    #[pymodule_init]
    fn init(module: &Bound<'_, PyModule>) -> PyResult<()> {
        module.add("__version__", env!("CARGO_PKG_VERSION"))
    }
}

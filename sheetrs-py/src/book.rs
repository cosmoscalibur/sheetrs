//! `Book`: a workbook that is a reference, with the operations that accumulate on it.

use crate::cells::{Block, Grids, Match};
use crate::convert::{json_to_py, py_to_json};
use crate::errors::{classified, format_error, invalid_range, pending, raise, to_py_err};
use crate::lint::{self, PyLintReport};
use crate::plan::{self, PlanError, Step, Target};
use crate::results::{PyInspection, PyReport, PySheets, PyStats, PyValueRanges, emit};
use pyo3::exceptions::{PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyByteArray, PyBytes, PyDict, PyList, PyString};
use serde_json::{Map, Value};
use sheetrs_compiler::memory::{self, BookCompile, Options};
use sheetrs_edit::{ColumnRange, RowRange, ValueScope};
use sheetrs_formula::DateSystem;
use sheetrs_io::Book;
use sheetrs_protocol::{Find, MajorDimension, Property, Replace, Scope, SearchIn};
use sheetrs_query::{BookQuery, FindOptions, Major, RangeOptions};
use sheetrs_stats::{BookStats, WorkbookStats};
use std::path::PathBuf;
use std::sync::{Arc, Mutex};

/// What the read that happens when a `Book` is opened found out: the statistics (which hold the
/// sheets, their states, the macros and the defined names) and the date system, or why the date
/// system cannot be told.
struct Opened {
    stats: Arc<WorkbookStats>,
    date_system: Result<DateSystem, String>,
}

/// A workbook: a file, its bytes or a binary stream, with the operations to make to it.
///
/// A `Book` is a reference: it reads the workbook once when it is opened, in a single streaming pass
/// that counts what it holds (the sheets, the used cells, the formulas, the names, the macros), and it
/// keeps only those numbers. The queries (`range`, `find`, `lint`) read what they need, and
/// the operations (`delete_sheets`, `delete_names`, `formula_to_value`, `delete_rows`,
/// `delete_columns`, `replace`, `calculate`) only accumulate: every one is added to the same
/// `Book`, which it returns so the calls chain, and each is written against the workbook as it is:
/// the order they are given in changes nothing, the planner runs them in the fewest passes over the
/// file in an order of its own (`explain()` says which). Nothing runs until `save`, the cut. To work
/// step by step, save between steps.
///
/// A query on a `Book` that has operations nothing has run raises `PendingOperationsError`: the
/// answer would be about a workbook that does not exist yet.
#[pyclass(name = "Book", module = "sheetrs")]
pub struct PyBook {
    source: Book,
    steps: Vec<Step>,
    report: Option<PyReport>,
    /// The binary stream the workbook was read from or written to: `save()` rewrites it.
    stream: Option<Py<PyAny>>,
    /// What the read at open found out; `None` until it is made (a `Book` that `save` returns reads
    /// the file it wrote the first time it is asked).
    opened: Mutex<Option<Arc<Opened>>>,
}

/// A query that cannot be answered: a range or sheet that is not valid is an `InvalidRangeError`, a
/// search the matcher does not reproduce a `ValueError`, and a workbook that cannot be read an `Error`.
fn query_error(error: sheetrs_query::Error) -> PyErr {
    match error {
        sheetrs_query::Error::Request(message) => invalid_range(message),
        sheetrs_query::Error::Read(e) => to_py_err(e),
        other => PyValueError::new_err(other.to_string()),
    }
}

fn plan_error(error: PlanError) -> PyErr {
    to_py_err(error)
}

/// Each target parsed as `T` (`ValueScope`, `RowRange`, `ColumnRange`): a malformed one, or one of
/// the wrong kind (rows deleted with a column range), is an `InvalidRangeError` where it is
/// written. `values` is one `str` or an iterable of them (a sweep gives its array); anything else
/// in it is a `TypeError`.
fn parse_all<T>(values: &Bound<'_, PyAny>, what: &str) -> PyResult<Vec<T>>
where
    T: std::str::FromStr,
    T::Err: std::fmt::Display,
{
    texts(values, what)?
        .iter()
        .map(|spec| {
            spec.parse::<T>()
                .map_err(|e| invalid_range(format!("invalid scope '{spec}': {e}")))
        })
        .collect()
}

/// One text or an iterable of them, as a list.
fn texts(value: &Bound<'_, PyAny>, what: &str) -> PyResult<Vec<String>> {
    if let Ok(one) = value.cast::<PyString>() {
        return Ok(vec![one.to_str()?.to_owned()]);
    }
    value
        .try_iter()
        .and_then(|iter| iter.map(|item| item?.extract::<String>()).collect())
        .map_err(|_| PyTypeError::new_err(format!("{what} is a str or an iterable of str")))
}

/// The `(text, replacement)` pairs of a `replace`: one `str` with its `replacement`, a mapping, or
/// an iterable of pairs. An empty mapping or iterable is no pair, and does nothing.
fn replacement_pairs(
    text: &Bound<'_, PyAny>,
    replacement: Option<&str>,
) -> PyResult<Vec<(String, String)>> {
    if let Ok(one) = text.cast::<PyString>() {
        let with = replacement.ok_or_else(|| {
            PyTypeError::new_err(
                "replace(text, replacement): with a str `text`, `replacement` is the str that takes its place",
            )
        })?;
        return Ok(vec![(one.to_str()?.to_owned(), with.to_owned())]);
    }
    if replacement.is_some() {
        return Err(PyTypeError::new_err(
            "replace: `replacement` goes with a str `text`; a mapping or a list of pairs carries its own",
        ));
    }
    let wrong = || {
        PyTypeError::new_err(
            "replace: `text` is a str, a mapping {text: replacement} or an iterable of (text, replacement) pairs of str",
        )
    };
    let items: Bound<'_, PyAny> = if text.hasattr("items")? {
        text.call_method0("items")?
    } else {
        text.clone()
    };
    let mut pairs = Vec::new();
    for item in items.try_iter().map_err(|_| wrong())? {
        let item = item?;
        // A str is an iterable of its characters: it is not a pair.
        if item.cast::<PyString>().is_ok() {
            return Err(wrong());
        }
        let mut parts = item.try_iter().map_err(|_| wrong())?;
        let (Some(a), Some(b), None) = (parts.next(), parts.next(), parts.next()) else {
            return Err(wrong());
        };
        pairs.push((
            a?.extract::<String>().map_err(|_| wrong())?,
            b?.extract::<String>().map_err(|_| wrong())?,
        ));
    }
    Ok(pairs)
}

fn properties_of(value: Option<&Bound<'_, PyAny>>) -> PyResult<Vec<Property>> {
    let Some(value) = value else {
        return Ok(vec![Property::Value]);
    };
    let mut out = Vec::new();
    for name in texts(value, "`properties`")? {
        let property = match name.as_str() {
            "value" => Property::Value,
            "formula" => Property::Formula,
            "number_format" => Property::NumberFormat,
            "display" => Property::Display,
            "comment" => Property::Comment,
            other => {
                return Err(PyValueError::new_err(format!(
                    "`properties` takes 'value', 'formula', 'number_format', 'display' and 'comment', not '{other}'"
                )));
            }
        };
        if !out.contains(&property) {
            out.push(property);
        }
    }
    if out.is_empty() {
        return Err(PyValueError::new_err("`properties` is empty"));
    }
    Ok(out)
}

/// The Python value of a property of a cell.
fn py_value(py: Python<'_>, value: Option<&Value>) -> PyResult<Py<PyAny>> {
    value.map_or_else(|| Ok(py.None()), |v| json_to_py(py, v))
}

/// The Python value of a text property of a cell: a `str`, or `None`.
fn py_text(py: Python<'_>, text: Option<&String>) -> PyResult<Py<PyAny>> {
    match text {
        Some(text) => Ok(text.as_str().into_pyobject(py)?.into_any().unbind()),
        None => Ok(py.None()),
    }
}

/// The grids of a block of the query as the lists of lists Python has.
fn block_to_py(
    py: Python<'_>,
    block: sheetrs_query::Block,
    major: &'static str,
) -> PyResult<Block> {
    fn lists<T>(
        py: Python<'_>,
        grid: Option<Vec<Vec<Option<T>>>>,
        convert: impl Fn(Python<'_>, Option<&T>) -> PyResult<Py<PyAny>>,
    ) -> PyResult<Option<Py<PyList>>> {
        let Some(grid) = grid else {
            return Ok(None);
        };
        let out = PyList::empty(py);
        for row in grid {
            let cells = row
                .iter()
                .map(|cell| convert(py, cell.as_ref()))
                .collect::<PyResult<Vec<_>>>()?;
            out.append(PyList::new(py, cells)?)?;
        }
        Ok(Some(out.unbind()))
    }
    Ok(Block::new(
        block.range,
        major,
        Grids {
            values: lists(py, block.values, py_value)?,
            formulas: lists(py, block.formulas, |py, t: Option<&String>| py_text(py, t))?,
            number_formats: lists(py, block.number_formats, |py, t: Option<&String>| {
                py_text(py, t)
            })?,
            displays: lists(py, block.displays, |py, t: Option<&String>| py_text(py, t))?,
            comments: lists(py, block.comments, |py, t: Option<&String>| py_text(py, t))?,
        },
    ))
}

/// The bytes a stream holds, read from its start.
fn read_stream(stream: &Bound<'_, PyAny>) -> PyResult<Vec<u8>> {
    if stream.hasattr("seek")? {
        stream.call_method1("seek", (0,))?;
    }
    let data = stream.call_method0("read")?;
    if let Ok(bytes) = data.cast::<PyBytes>() {
        return Ok(bytes.as_bytes().to_vec());
    }
    if let Ok(bytes) = data.cast::<PyByteArray>() {
        return Ok(bytes.to_vec());
    }
    Err(PyTypeError::new_err(
        "a stream gives a workbook as bytes: open the file in binary mode (`open(path, \"rb\")`)",
    ))
}

/// Replaces what a stream holds with `bytes`: from its start, and cut where they end.
fn rewrite_stream(py: Python<'_>, stream: &Bound<'_, PyAny>, bytes: &[u8]) -> PyResult<()> {
    if stream.hasattr("seek")? {
        stream.call_method1("seek", (0,))?;
    }
    stream.call_method1("write", (PyBytes::new(py, bytes),))?;
    if stream.hasattr("truncate")? {
        stream.call_method0("truncate")?;
    }
    Ok(())
}

/// Whether the argument is a binary stream to write to (it has `write`) and not a path.
fn is_stream(value: &Bound<'_, PyAny>) -> PyResult<bool> {
    value.hasattr("write")
}

/// Writes `bytes` where `target` says: a stream gets them, a path gets a file written beside it and
/// renamed.
fn write_to_target(py: Python<'_>, target: &Bound<'_, PyAny>, bytes: &[u8]) -> PyResult<()> {
    if is_stream(target)? {
        target.call_method1("write", (PyBytes::new(py, bytes),))?;
        return Ok(());
    }
    let path = target
        .extract::<PathBuf>()
        .map_err(|_| PyTypeError::new_err("a destination is a path or a binary stream"))?;
    plan::write_atomic(&path, bytes).map_err(PyErr::from)
}

impl PyBook {
    fn from_source(source: Book) -> Self {
        Self {
            source,
            steps: Vec::new(),
            report: None,
            stream: None,
            opened: Mutex::new(None),
        }
    }

    /// What the read at open found out, made now if it was not. The lock is not held while the
    /// workbook is read: two threads that ask at once both read it, and the first answer stays.
    fn opened(&self, py: Python<'_>) -> PyResult<Arc<Opened>> {
        let poisoned = || raise("Error", "the Book is poisoned");
        if let Some(opened) = self.opened.lock().map_err(|_| poisoned())?.as_ref() {
            return Ok(Arc::clone(opened));
        }
        let source = &self.source;
        let read = py.detach(|| -> PyResult<Arc<Opened>> {
            let stats = source.stats().map_err(to_py_err)?;
            let date_system = source.date_system().map_err(|e| format!("{e:#}"));
            Ok(Arc::new(Opened {
                stats: Arc::new(stats),
                date_system,
            }))
        })?;
        let mut slot = self.opened.lock().map_err(|_| poisoned())?;
        Ok(Arc::clone(slot.get_or_insert(read)))
    }

    /// What the read at open found out, for a query: it is about the workbook as the file has it,
    /// so it cannot be asked while operations are pending.
    fn described(&self, py: Python<'_>) -> PyResult<Arc<Opened>> {
        self.queryable()?;
        self.opened(py)
    }

    /// Forgets what the read at open found out: the workbook changed under this `Book`.
    fn forget_opened(&self) {
        if let Ok(mut slot) = self.opened.lock() {
            *slot = None;
        }
    }

    /// A query is about the workbook as the file has it: it cannot run while operations are pending.
    fn queryable(&self) -> PyResult<()> {
        if self.steps.is_empty() {
            Ok(())
        } else {
            Err(pending(self.steps.len()))
        }
    }

    fn push(mut slf: PyRefMut<'_, Self>, step: Step) -> PyRefMut<'_, Self> {
        slf.steps.push(step);
        slf
    }

    /// A call with no target does nothing: it adds no operation (so it costs no pass).
    fn push_unless_empty(slf: PyRefMut<'_, Self>, empty: bool, step: Step) -> PyRefMut<'_, Self> {
        if empty { slf } else { Self::push(slf, step) }
    }
}

#[pymethods]
impl PyBook {
    /// A workbook: the path of a file (`str` or `os.PathLike`), its bytes (`bytes` or `bytearray`) or
    /// a binary stream that reads and seeks (`io.BytesIO`, a file opened with `"rb"`). The format (XLSX,
    /// XLSM or ODS) is the one of the content, whatever the extension says.
    ///
    /// The workbook is read once, here, to count what it holds (see the properties); a file is never
    /// modified except by `save()` without a destination, which also rewrites a stream the `Book` was
    /// made from.
    ///
    /// Raises `FileNotFoundError` for a path that is not a file and `FormatError` for a workbook
    /// that is not an XLSX, XLSM or ODS one.
    #[new]
    fn new(py: Python<'_>, source: &Bound<'_, PyAny>) -> PyResult<Self> {
        let mut book = if let Ok(bytes) = source.cast::<PyBytes>() {
            Self::from_source(Book::from_bytes(bytes.as_bytes().to_vec()).map_err(format_error)?)
        } else if let Ok(bytes) = source.cast::<PyByteArray>() {
            Self::from_source(Book::from_bytes(bytes.to_vec()).map_err(format_error)?)
        } else if source.hasattr("read")? {
            let data = read_stream(source)?;
            let mut book = Self::from_source(Book::from_bytes(data).map_err(format_error)?);
            book.stream = Some(source.clone().unbind());
            book
        } else {
            let path = source.extract::<PathBuf>().map_err(|_| {
                PyTypeError::new_err(
                    "a workbook is a path (str or os.PathLike), bytes or a binary stream",
                )
            })?;
            if !path.is_file() {
                return Err(std::io::Error::from(std::io::ErrorKind::NotFound).into());
            }
            let opened = Book::from_file(&path).map_err(|error| match error {
                sheetrs_io::Error::Open { .. } => to_py_err(error),
                other => format_error(other),
            })?;
            Self::from_source(opened)
        };
        book.opened(py)?;
        book.report = None;
        Ok(book)
    }

    /// `"xlsx"`, `"xlsm"` or `"ods"`.
    #[getter]
    fn format(&self) -> &'static str {
        self.source.format().extension()
    }

    /// The date of serial 0, as a `datetime.date`: `1899-12-30` (the default of `LibreOffice` and of an
    /// XLSX), `1900-01-01` or `1904-01-01`. A workbook whose date system is none of the three raises
    /// `UnsupportedDateSystemError`.
    #[getter]
    fn date_system<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let opened = self.described(py)?;
        match &opened.date_system {
            Ok(system) => {
                let (year, month, day) = system.null_date();
                py.import("datetime")?
                    .getattr("date")?
                    .call1((year, month, day))
            }
            Err(message) => Err(classified(message.clone())),
        }
    }

    /// Whether the workbook carries macros (found, never run).
    #[getter]
    fn has_macros(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.described(py)?.stats.has_macros)
    }

    /// The names of the sheets, in order. Raises `PendingOperationsError` with operations pending.
    #[getter]
    fn sheet_names(&self, py: Python<'_>) -> PyResult<Vec<String>> {
        Ok(self
            .described(py)?
            .stats
            .sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect())
    }

    /// The sheets, by name (`book.sheets["Data"]`) or by position, each with its `name`, `used_range`,
    /// `state` and counts.
    #[getter]
    fn sheets(&self, py: Python<'_>) -> PyResult<PySheets> {
        Ok(PySheets::of(&self.described(py)?.stats))
    }

    /// The defined names and what each refers to, as the file writes it (`{"Rate": "Data!$B$2"}`),
    /// ordered by name.
    #[getter]
    fn defined_names<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let dict = PyDict::new(py);
        for (name, refers_to) in &self.described(py)?.stats.defined_names {
            dict.set_item(name, refers_to)?;
        }
        Ok(dict)
    }

    /// The statistics of the workbook, as `sheetrs stats` reports them: the totals (also on the
    /// `Book`, as `total_*`) and the sheets.
    #[getter]
    fn stats(&self, py: Python<'_>) -> PyResult<PyStats> {
        Ok(PyStats::of(Arc::clone(&self.described(py)?.stats)))
    }

    #[getter]
    fn total_sheets(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(self.described(py)?.stats.total_sheets)
    }

    #[getter]
    fn total_named_ranges(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(self.described(py)?.stats.total_named_ranges)
    }

    /// The size of the file in bytes; 0 for a workbook given as bytes or as a stream.
    #[getter]
    fn total_file_size(&self, py: Python<'_>) -> PyResult<u64> {
        Ok(self.described(py)?.stats.total_file_size)
    }

    #[getter]
    fn total_used_cells(&self, py: Python<'_>) -> PyResult<u64> {
        Ok(self.described(py)?.stats.total_used_cells)
    }

    #[getter]
    fn total_formulas(&self, py: Python<'_>) -> PyResult<u64> {
        Ok(self.described(py)?.stats.total_formulas)
    }

    #[getter]
    fn total_errors(&self, py: Python<'_>) -> PyResult<u64> {
        Ok(self.described(py)?.stats.total_errors)
    }

    #[getter]
    fn total_empty_cells(&self, py: Python<'_>) -> PyResult<u64> {
        Ok(self.described(py)?.stats.total_empty_cells)
    }

    #[getter]
    fn total_comments(&self, py: Python<'_>) -> PyResult<u64> {
        Ok(self.described(py)?.stats.total_comments)
    }

    #[getter]
    fn total_charts(&self, py: Python<'_>) -> PyResult<u64> {
        Ok(self.described(py)?.stats.total_charts)
    }

    #[getter]
    fn total_tables(&self, py: Python<'_>) -> PyResult<u64> {
        Ok(self.described(py)?.stats.total_tables)
    }

    #[getter]
    fn total_xml_size(&self, py: Python<'_>) -> PyResult<u64> {
        Ok(self.described(py)?.stats.total_xml_size)
    }

    /// What the last `save` or `to_bytes` did besides writing the workbook, as a `Report` (the cells a
    /// `replace` changed, the cells it left, and the `diagnostics`); `None` for a `Book` nothing wrote.
    #[getter]
    fn report(&self) -> Option<PyReport> {
        self.report.clone()
    }

    // ---- the operations: they accumulate, and `save` runs them -------------------------------

    /// Delete sheets by name (*Delete Sheet*): one name or an iterable of them. An empty iterable
    /// does nothing.
    fn delete_sheets<'py>(
        slf: PyRefMut<'py, Self>,
        names: &Bound<'_, PyAny>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let names = texts(names, "`names`")?;
        Ok(Self::push_unless_empty(
            slf,
            names.is_empty(),
            Step::DeleteSheets(names),
        ))
    }

    /// Delete defined names (and database ranges): *Manage Names ▸ Delete*. One name or an
    /// iterable of them.
    fn delete_names<'py>(
        slf: PyRefMut<'py, Self>,
        names: &Bound<'_, PyAny>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let names = texts(names, "`names`")?;
        Ok(Self::push_unless_empty(
            slf,
            names.is_empty(),
            Step::DeleteNames(names),
        ))
    }

    /// Replace the formulas of the scopes by the values they have (*Data ▸ Calculate ▸ Formula to
    /// Value*). A scope is `"Sheet"`, `"Sheet!B:D"` (columns) or `"Sheet!3:10"` (rows); one or an
    /// iterable of them.
    fn formula_to_value<'py>(
        slf: PyRefMut<'py, Self>,
        scopes: &Bound<'_, PyAny>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let scopes = parse_all::<ValueScope>(scopes, "`scopes`")?;
        Ok(Self::push_unless_empty(
            slf,
            scopes.is_empty(),
            Step::FormulaToValue(scopes),
        ))
    }

    /// Delete rows (`"Sheet!3:10"`, or `"Sheet!3"`): one or an iterable of them. References to them
    /// become `#REF!` and the ranges around them shrink. The ranges of one call are simultaneous.
    fn delete_rows<'py>(
        slf: PyRefMut<'py, Self>,
        rows: &Bound<'_, PyAny>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let rows = parse_all::<RowRange>(rows, "`rows`")?;
        Ok(Self::push_unless_empty(
            slf,
            rows.is_empty(),
            Step::DeleteRows(rows),
        ))
    }

    /// Delete columns (`"Sheet!B:D"`, or `"Sheet!B"`): one or an iterable of them. The ranges of
    /// one call are simultaneous.
    fn delete_columns<'py>(
        slf: PyRefMut<'py, Self>,
        columns: &Bound<'_, PyAny>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let columns = parse_all::<ColumnRange>(columns, "`columns`")?;
        Ok(Self::push_unless_empty(
            slf,
            columns.is_empty(),
            Step::DeleteColumns(columns),
        ))
    }

    /// Find `text` and replace it with `replacement`, as Find & Replace of `LibreOffice` does (it
    /// reaches formula cells too). The options are those of `find`; `scope` is a sheet, a range, or
    /// an iterable of them (the whole workbook when omitted). What it did is in `report` of the
    /// `Book` that `save` returns.
    ///
    /// A sweep gives all its replacements at once: `text` is then a mapping `{text: replacement}`
    /// or an iterable of `(text, replacement)` pairs, with no `replacement`. Each pair is an
    /// operation of the plan, in order, that sees the workbook the pair before it left, exactly as
    /// if `replace` were called once per pair; the options and the scope are shared. The order
    /// matters (`{"a": "b", "b": "c"}` turns `a` into `c`).
    #[pyo3(signature = (text, replacement=None, *, regex=false, wildcards=false, match_case=false,
        entire_cells=false, search_in="values", formatted=false, filtered=false, scope=None))]
    #[allow(clippy::too_many_arguments, clippy::fn_params_excessive_bools)]
    fn replace<'py>(
        mut slf: PyRefMut<'py, Self>,
        text: &Bound<'py, PyAny>,
        replacement: Option<&str>,
        regex: bool,
        wildcards: bool,
        match_case: bool,
        entire_cells: bool,
        search_in: &str,
        formatted: bool,
        filtered: bool,
        scope: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let pairs = replacement_pairs(text, replacement)?;
        let scopes: Vec<Scope> = match scope {
            None => Vec::new(),
            Some(value) => texts(value, "`scope`")?
                .iter()
                .map(|s| Scope::from(s.as_str()))
                .collect(),
        };
        // Every pair is checked before any is added: a mistake in the third leaves the plan as it was.
        let mut steps = Vec::with_capacity(pairs.len());
        for (text, with) in pairs {
            let mut find = find_of(
                &text,
                regex,
                wildcards,
                match_case,
                entire_cells,
                search_in,
                true,
                formatted,
                filtered,
            )?;
            find.include_formulas = true;
            steps.push(Step::Replace {
                scopes: scopes.clone(),
                replace: Replace { find, with },
            });
        }
        slf.steps.extend(steps);
        Ok(slf)
    }

    /// Calculate with the engine, which reproduces `LibreOffice` (*Data ▸ Calculate*): an operation
    /// like the others, not a query. The engine writes `inputs` into their cells first when they are
    /// given (`{"Data!B2": 1500}`) and calculates every formula of the workbook as the operations
    /// before it leave it, and what comes after addresses the result. `save` writes the workbook with
    /// the results and without the mark to recalculate it when it is opened, so no other program has
    /// to open it for its results to be up to date. It returns the `Book`.
    ///
    /// `ranges` (a range, or a list of them: `["Summary!B2:B5"]`) are the ranges whose values
    /// `value_ranges()` gives and `save(json=...)` writes; without them, the outputs the workbook declares
    /// or, with none declared, its formula cells. `major` is `"rows"` (the default) or `"columns"`:
    /// how each range is listed, as the `majorDimension` of the request.
    ///
    /// Where the engine cannot calculate the workbook (a function it does not have, a reference it
    /// does not read) `save` still writes it, marked to be recalculated when it is opened, and issues
    /// a `RecalculationWarning`: the recalculation that is needed will be done by the spreadsheet
    /// program that opens the file later. With `to_dict()` or `save(json=...)` there is no answer to
    /// give, and it is an error.
    #[pyo3(signature = (inputs=None, ranges=None, *, major="rows"))]
    fn calculate<'py>(
        mut slf: PyRefMut<'py, Self>,
        py: Python<'py>,
        inputs: Option<&Bound<'py, PyDict>>,
        ranges: Option<&Bound<'py, PyAny>>,
        major: &str,
    ) -> PyResult<Bound<'py, PyAny>> {
        let inputs: Map<String, Value> = match inputs {
            Some(dict) => dict
                .iter()
                .map(|(key, value)| Ok((key.extract::<String>()?, py_to_json(&value)?)))
                .collect::<PyResult<Map<String, Value>>>()?,
            None => Map::new(),
        };
        let ranges = ranges.map(|o| texts(o, "`ranges`")).transpose()?;
        let major = match major {
            "rows" => MajorDimension::Rows,
            "columns" => MajorDimension::Columns,
            other => {
                return Err(PyValueError::new_err(format!(
                    "`major` is 'rows' or 'columns', not '{other}'"
                )));
            }
        };
        slf.steps.push(Step::Calculate {
            inputs,
            ranges,
            major,
        });
        Ok(slf.into_pyobject(py)?.into_any())
    }

    /// The value ranges of the calculation, as a `ValueRanges` (a mapping from the name of each one to
    /// a `Block`; `to_dict()` gives `{name: list[list]}`): the ranges `calculate(ranges=...)` asked
    /// for, by the text they were asked with, or the outputs of the workbook by their names.
    /// It is a terminal like `save`: it runs the operations (the workbook goes through memory and is
    /// written nowhere), so it is allowed with operations pending. Raises `Error` when there is no
    /// `calculate` step, and `UnavailableFunctionError` when the engine cannot calculate the
    /// workbook.
    fn value_ranges(&self, py: Python<'_>) -> PyResult<PyValueRanges> {
        if !plan::has_calculate(&self.steps) {
            return Err(raise(
                "Error",
                "there is no calculate step: value_ranges() gives the value ranges of a calculation, so \
                 call calculate() first",
            ));
        }
        let source = self.source.clone();
        let steps = self.steps.clone();
        let done = py
            .detach(|| plan::run(&source, &steps, &Target::Nowhere, true))
            .map_err(plan_error)?;
        let asked = plan::asked_ranges(&steps);
        let major = plan::asked_major(&steps);
        let value_ranges = done
            .report
            .value_ranges
            .ok_or_else(|| raise("Error", "the calculation gave no value ranges"))?;
        Ok(PyValueRanges::new(value_ranges, asked.as_deref(), major))
    }

    /// What `save` will do: the passes over the file that run the operations, one operation per
    /// line, and the calculations in their place. The operations are not run in the order they were
    /// given: each is written against the workbook as it is, and the planner orders them.
    fn explain(&self) -> String {
        plan::explain(&self.source, &self.steps)
    }

    /// The cut: run the operations and write the workbook.
    ///
    /// `save(dest)` writes to `dest`, a path (which must not be the file being edited) or a binary
    /// stream (`io.BytesIO()`), and returns a `Book` on what it wrote; this `Book` keeps its
    /// operations, so it can be saved elsewhere too. `save()` without a destination replaces the
    /// original: a file atomically (a failure leaves it as it was), or the stream the `Book` was
    /// made from, rewritten from its start. It returns this `Book`, which then has no operations: it is
    /// the workbook now. A `Book` opened from bytes has no original to replace.
    ///
    /// `json` is a path or a binary stream: the value ranges of the `calculate` step are written there as
    /// JSON (the answer of `sheetc calculate`) besides the workbook. It is an error when there is no
    /// `calculate` step, or when the engine cannot calculate the workbook.
    ///
    /// The editor runs the operations in the fewest passes over the file, streaming. A workbook
    /// the operations leave stale is calculated again by the engine, so it is written with fresh
    /// results; where the engine cannot, it is written marked so that the application that opens it
    /// recalculates every formula, and a `RecalculationWarning` says so (also logged by `sheetrs`).
    /// Raises `OSError` for a file that cannot be read or written and `sheetrs.Error` when the
    /// workbook cannot be edited (including when there are no operations).
    #[pyo3(signature = (dest=None, *, json=None))]
    fn save<'py>(
        mut slf: PyRefMut<'py, Self>,
        py: Python<'py>,
        dest: Option<&Bound<'py, PyAny>>,
        json: Option<&Bound<'py, PyAny>>,
    ) -> PyResult<Bound<'py, PyAny>> {
        if json.is_some() && !plan::has_calculate(&slf.steps) {
            return Err(raise(
                "Error",
                "there is no calculate step: save(json=...) writes the value ranges of a calculation, so \
                 call calculate() first",
            ));
        }
        let source = slf.source.clone();
        let steps = slf.steps.clone();
        let stream_dest = match dest {
            Some(dest) if is_stream(dest)? => Some(dest.clone()),
            _ => None,
        };
        let path_dest: Option<PathBuf> =
            match (dest, &stream_dest) {
                (Some(dest), None) => Some(dest.extract::<PathBuf>().map_err(|_| {
                    PyTypeError::new_err("a destination is a path or a binary stream")
                })?),
                _ => None,
            };
        // The stream the workbook came from takes the result of `save()` with no destination.
        let source_stream = if dest.is_none() {
            slf.stream.as_ref().map(|s| s.bind(py).clone())
        } else {
            None
        };
        let to_memory = stream_dest.is_some() || source_stream.is_some();
        // Ranges named in `calculate(ranges=...)` are an answer to keep, as `json=` asks for it.
        let keep_answer = json.is_some() || plan::asked_ranges(&steps).is_some();
        let done = py
            .detach(|| {
                let target = match (&path_dest, to_memory) {
                    (_, true) => Target::Bytes,
                    (Some(path), false) => Target::Path(path),
                    (None, false) => Target::InPlace,
                };
                plan::run(&source, &steps, &target, keep_answer)
            })
            .map_err(plan_error)?;
        if let Some(json) = json {
            let value_ranges = done
                .report
                .value_ranges
                .as_ref()
                .ok_or_else(|| raise("Error", "the calculation gave no value ranges"))?;
            let mut text = Vec::new();
            value_ranges
                .write_json(&mut text)
                .map_err(|e| raise("Error", format!("cannot write the value ranges: {e}")))?;
            write_to_target(py, json, &text)?;
        }
        let report = PyReport::of(&done.report);
        let diagnostics = done.report.diagnostics.clone();
        let result = if let Some(stream) = &stream_dest {
            let bytes = done
                .bytes
                .ok_or_else(|| raise("Error", "the workbook was not written"))?;
            write_to_target(py, stream, &bytes)?;
            let mut saved = Self::from_source(Book::from_bytes(bytes).map_err(to_py_err)?);
            saved.stream = Some(stream.clone().unbind());
            saved.report = Some(report);
            Bound::new(py, saved)?.into_any()
        } else if let Some(path) = path_dest {
            let source = Book::from_file(path).map_err(to_py_err)?;
            let mut saved = Self::from_source(source);
            saved.report = Some(report);
            Bound::new(py, saved)?.into_any()
        } else {
            if let Some(stream) = &source_stream {
                let bytes = done
                    .bytes
                    .ok_or_else(|| raise("Error", "the workbook was not written"))?;
                rewrite_stream(py, stream, &bytes)?;
                slf.source = Book::from_bytes(bytes).map_err(to_py_err)?;
            }
            slf.steps.clear();
            slf.report = Some(report);
            slf.forget_opened();
            slf.into_pyobject(py)?.into_any()
        };
        emit(py, &diagnostics)?;
        Ok(result)
    }

    /// The workbook the operations make, as bytes, in the format of the source (`xlsx`, `xlsm` or
    /// `ods`): the cut like `save`, but to memory, so no file is made. It runs the accumulated
    /// operations in the fewest passes (`explain()` says which), streaming, and the bytes are the
    /// ones `save(path)` would write for the same plan.
    ///
    /// Like `save(dest)`, it leaves this `Book` with its operations (the plan is not consumed), so
    /// it can be written again elsewhere; `report` says what this run did, and a
    /// `RecalculationWarning` is issued where `save` would issue it. With no operations pending it
    /// returns the bytes of the workbook as they are (the file is read, or the bytes the `Book` was
    /// made from).
    ///
    /// Raises `OSError` for a file that cannot be read when there are no operations, and
    /// `sheetrs.Error` when the workbook cannot be read or edited to run them.
    fn to_bytes<'py>(
        mut slf: PyRefMut<'py, Self>,
        py: Python<'py>,
    ) -> PyResult<Bound<'py, PyBytes>> {
        let source = slf.source.clone();
        if slf.steps.is_empty() {
            let bytes = py
                .detach(|| match (source.path(), source.bytes()) {
                    (Some(path), _) => std::fs::read(path),
                    (None, Some(data)) => Ok(data.to_vec()),
                    (None, None) => unreachable!("a book is a file or bytes"),
                })
                .map_err(PyErr::from)?;
            return Ok(PyBytes::new(py, &bytes));
        }
        let steps = slf.steps.clone();
        let done = py
            .detach(|| plan::run(&source, &steps, &Target::Bytes, false))
            .map_err(plan_error)?;
        let bytes = done
            .bytes
            .ok_or_else(|| raise("Error", "the workbook was not written"))?;
        slf.report = Some(PyReport::of(&done.report));
        emit(py, &done.report.diagnostics)?;
        Ok(PyBytes::new(py, &bytes))
    }

    // ---- the queries: they read what they need -------------------------------------------------

    /// The cells of one or several ranges, read in a single pass over the file.
    ///
    /// `ranges` is a range (`"Data!A1:C10"`) or a list of them, in any sheets; a defined name that
    /// is a range is accepted. It returns `{range: Block}` in the order asked. `properties` are
    /// `"value"` (the default), `"formula"`, `"number_format"`, `"display"` and `"comment"`: only
    /// what is asked is read. `major="columns"` gives lists of columns. Numbers are floats (ints when
    /// exact); a date is a number unless `dates=True`, which gives `datetime.date`,
    /// `datetime.time` or `datetime.datetime` to the cells whose format is a date or a time, by the
    /// date system of the file; an error is a `CellError`.
    ///
    /// Raises `InvalidRangeError` for a range or a sheet that does not exist and
    /// `PendingOperationsError` with operations pending.
    #[pyo3(signature = (ranges, properties=None, *, major="rows", dates=false))]
    fn range<'py>(
        &self,
        py: Python<'py>,
        ranges: &Bound<'py, PyAny>,
        properties: Option<&Bound<'py, PyAny>>,
        major: &str,
        dates: bool,
    ) -> PyResult<Bound<'py, PyDict>> {
        self.queryable()?;
        let major_name: &'static str = match major {
            "rows" => "rows",
            "columns" => "columns",
            other => {
                return Err(PyValueError::new_err(format!(
                    "`major` is 'rows' or 'columns', not '{other}'"
                )));
            }
        };
        let options = RangeOptions {
            properties: properties_of(properties)?,
            major: if major_name == "columns" {
                Major::Columns
            } else {
                Major::Rows
            },
            dates,
        };
        let asked = texts(ranges, "`ranges`")?;
        let source = &self.source;
        let blocks = py
            .detach(|| source.range(&asked, &options))
            .map_err(query_error)?;
        let out = PyDict::new(py);
        for block in blocks {
            let key = block.range.clone();
            out.set_item(&key, Bound::new(py, block_to_py(py, block, major_name)?)?)?;
        }
        Ok(out)
    }

    /// Find the cells that match `text`, as Find & Replace of `LibreOffice` does: a list of `Match`
    /// (sheet, cell, row, col and the `properties` asked for, `"value"` by default).
    ///
    /// `regex` (ICU regular expression) and `wildcards` (`*`, `?`, `\`) are exclusive; `match_case`;
    /// `entire_cells` (the text must be the whole content); `search_in` is `"values"` (the default),
    /// `"formulas"` or `"comments"`; `include_formulas` also searches the result of formula cells
    /// when searching values; `formatted` searches the text the cell shows; `filtered` reaches the
    /// rows an autofilter hides; `scope` is a sheet, a range or a list of them (the whole workbook
    /// when omitted); `by` orders the matches by `"rows"` (the default) or `"columns"`; `dates` as in
    /// `range`.
    #[pyo3(signature = (text, *, regex=false, wildcards=false, match_case=false, entire_cells=false,
        search_in="values", include_formulas=false, formatted=false, filtered=false, scope=None,
        properties=None, by="rows", dates=false))]
    #[allow(clippy::too_many_arguments, clippy::fn_params_excessive_bools)]
    fn find<'py>(
        &self,
        py: Python<'py>,
        text: &str,
        regex: bool,
        wildcards: bool,
        match_case: bool,
        entire_cells: bool,
        search_in: &str,
        include_formulas: bool,
        formatted: bool,
        filtered: bool,
        scope: Option<&Bound<'py, PyAny>>,
        properties: Option<&Bound<'py, PyAny>>,
        by: &str,
        dates: bool,
    ) -> PyResult<Bound<'py, PyList>> {
        self.queryable()?;
        if by != "rows" && by != "columns" {
            return Err(PyValueError::new_err(format!(
                "`by` is 'rows' or 'columns', not '{by}'"
            )));
        }
        let find = find_of(
            text,
            regex,
            wildcards,
            match_case,
            entire_cells,
            search_in,
            include_formulas,
            formatted,
            filtered,
        )?;
        let scopes = match scope {
            None => Vec::new(),
            Some(value) => texts(value, "`scope`")?
                .iter()
                .map(|s| Scope::from(s.as_str()))
                .collect(),
        };
        let options = FindOptions {
            scope: scopes,
            properties: properties_of(properties)?,
            by: if by == "columns" {
                Major::Columns
            } else {
                Major::Rows
            },
            dates,
        };
        let source = &self.source;
        let cells = py
            .detach(|| source.find(&find, &options))
            .map_err(query_error)?;
        let out = PyList::empty(py);
        for cell in cells {
            let value = py_value(py, cell.value.as_ref())?;
            out.append(Bound::new(
                py,
                Match {
                    sheet: cell.sheet,
                    row: cell.row,
                    col: cell.col,
                    value,
                    formula: cell.formula,
                    number_format: cell.number_format,
                    display: cell.display,
                    comment: cell.comment,
                },
            )?)?;
        }
        Ok(out)
    }

    /// Lint the workbook with the configuration of the file `config_path` (the default rules when
    /// it is `None`; an unknown rule, category or parameter is a `ValueError`). It returns a
    /// `LintReport`: a sequence of `Violation`, ordered by scope (book, sheet, cell) and rule, that
    /// writes itself as `sheetrs lint` does (`to_text()`, `to_json()`, `to_sarif()`, `to_dict()`).
    #[pyo3(signature = (config_path=None))]
    #[allow(clippy::needless_pass_by_value)] // pyo3 hands over an owned argument
    fn lint(
        &self,
        py: Python<'_>,
        config_path: Option<std::borrow::Cow<'_, std::path::Path>>,
    ) -> PyResult<PyLintReport> {
        self.queryable()?;
        lint::lint(py, &self.source, config_path.as_deref())
    }

    /// What `sheetc inspect` says about the workbook: its sheets and date system, the cells and
    /// formulas, the cycles, the functions it uses and which of them the engine lacks. It compiles the
    /// workbook in memory and calculates nothing.
    fn inspect(&self, py: Python<'_>) -> PyResult<PyInspection> {
        self.queryable()?;
        let source = &self.source;
        let inspection = py
            .detach(|| source.inspect(&Options::default()))
            .map_err(to_py_err)?;
        Ok(PyInspection::of(inspection))
    }

    fn __repr__(&self) -> String {
        let origin = self.source.path().map_or_else(
            || {
                if self.stream.is_some() {
                    format!("{} stream", self.source.format().extension())
                } else {
                    format!("{} bytes", self.source.format().extension())
                }
            },
            |path| path.display().to_string(),
        );
        match self.steps.len() {
            0 => format!("Book({origin:?})"),
            1 => format!("Book({origin:?}, 1 pending operation)"),
            n => format!("Book({origin:?}, {n} pending operations)"),
        }
    }
}

/// The search of the arguments, as the protocol has it.
#[allow(clippy::too_many_arguments, clippy::fn_params_excessive_bools)]
fn find_of(
    text: &str,
    regex: bool,
    wildcards: bool,
    match_case: bool,
    entire_cells: bool,
    search_in: &str,
    include_formulas: bool,
    formatted: bool,
    filtered: bool,
) -> PyResult<Find> {
    use sheetrs_protocol::Mode as Matching;
    let mode = match (regex, wildcards) {
        (true, true) => {
            return Err(PyValueError::new_err(
                "`regex` and `wildcards` cannot both be on",
            ));
        }
        (true, false) => Matching::Regex,
        (false, true) => Matching::Wildcards,
        (false, false) => Matching::Literal,
    };
    let search_in = match search_in {
        "values" => SearchIn::Value,
        "formulas" => SearchIn::Formula,
        "comments" => SearchIn::Comment,
        other => {
            return Err(PyValueError::new_err(format!(
                "`search_in` is 'values', 'formulas' or 'comments', not '{other}'"
            )));
        }
    };
    Ok(Find {
        text: text.to_owned(),
        mode,
        match_case,
        entire_cells,
        search_in,
        include_formulas,
        formatted,
        filtered,
    })
}

/// The value of one formula, with the regional settings of `locale` (a BCP-47 tag, `"en"` by
/// default).
pub(crate) fn evaluate(py: Python<'_>, formula: &str, locale: Option<&str>) -> PyResult<Py<PyAny>> {
    let value = py
        .detach(|| memory::evaluate(formula, locale, &Options::default()))
        .map_err(to_py_err)?;
    json_to_py(py, &value)
}

//! `Book`: a workbook that is a reference, with the operations that accumulate on it.

use crate::cells::{Block, Grids, Match};
use crate::convert::{json_to_py, py_to_json};
use crate::errors::{format_error, invalid_range, pending, raise, to_py_err};
use crate::lint::{self, PyLintReport};
use crate::plan::{self, PlanError, Step, Target};
use crate::results::{PyInspection, PyReport, PyStats, PyValueRanges, emit};
use pyo3::exceptions::{PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyByteArray, PyBytes, PyDict, PyList, PyString};
use serde_json::{Map, Value};
use sheetrs_compiler::memory::{self, BookCompile, Options};
use sheetrs_edit::{ColumnRange, RowRange, ValueScope, VisibilityChange};
use sheetrs_io::Book;
use sheetrs_protocol::{Find, MajorDimension, Property, Replace, Scope, SearchIn};
use sheetrs_query::{BookQuery, FindOptions, Major, RangeOptions};
use sheetrs_stats::BookStats;
use std::path::PathBuf;
use std::sync::Arc;

/// A workbook: a file, its bytes or a binary stream, with the operations to make to it.
///
/// A `Book` is a reference: it reads nothing when it is made, and keeps nothing of the workbook. The queries (`range`, `find`, `lint`) read what they need, and
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

    fn rows_visibility<'py>(
        slf: PyRefMut<'py, Self>,
        rows: &Bound<'_, PyAny>,
        visible: bool,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let changes: Vec<VisibilityChange> = parse_all::<RowRange>(rows, "`rows`")?
            .into_iter()
            .map(|range| VisibilityChange::Rows { range, visible })
            .collect();
        Ok(Self::push_unless_empty(
            slf,
            changes.is_empty(),
            Step::SetVisibility(changes),
        ))
    }

    fn columns_visibility<'py>(
        slf: PyRefMut<'py, Self>,
        columns: &Bound<'_, PyAny>,
        visible: bool,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let changes: Vec<VisibilityChange> = parse_all::<ColumnRange>(columns, "`columns`")?
            .into_iter()
            .map(|range| VisibilityChange::Columns { range, visible })
            .collect();
        Ok(Self::push_unless_empty(
            slf,
            changes.is_empty(),
            Step::SetVisibility(changes),
        ))
    }

    fn sheets_visibility<'py>(
        slf: PyRefMut<'py, Self>,
        names: &Bound<'_, PyAny>,
        visible: bool,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let changes: Vec<VisibilityChange> = texts(names, "`names`")?
            .into_iter()
            .map(|name| VisibilityChange::Sheet { name, visible })
            .collect();
        Ok(Self::push_unless_empty(
            slf,
            changes.is_empty(),
            Step::SetVisibility(changes),
        ))
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
    /// The workbook is not read here: the properties that describe it read it when they are first asked
    /// (once), and the queries read what they need. A file is never
    /// modified except by `save()` without a destination, which also rewrites a stream the `Book` was
    /// made from.
    ///
    /// `locale` is the locale of the workbook (a tag such as `"es-CO"`): the regional settings that give
    /// a text to a number (`TEXT`, the conversion of a number to text, `display`) and read a number from
    /// a text, in `calculate`, in the recalculation that follows an edit and in the queries. Without
    /// it the workbook, an XLSX or an ODS, has the locale of the system (`LC_ALL`, `LC_CTYPE` or
    /// `LANG`, read when it is needed; `en-US` when there is none). The `locale` of a single
    /// call (`evaluate`) wins over it.
    ///
    /// `now` (the serial number that `NOW()` and `TODAY()` give) and `seed` (the seed of `RAND()` and
    /// its relatives, a non-negative integer) fix the volatile functions in every calculation of the
    /// workbook: `calculate`, the recalculation that follows an edit and `inspect`; they are the
    /// `--now` and `--seed` of `sheetc`. Without them the engine reads the clock. The `now` and
    /// `seed` of a single call (`calculate(now=, seed=)`) win over them. Distinct seeds give distinct
    /// streams; `RAND` of LibreOffice cannot be seeded, so `seed` is an extension of SheetRS.
    ///
    /// Raises `FileNotFoundError` for a path that is not a file, `FormatError` for a workbook
    /// that is not an XLSX, XLSM or ODS one and `ValueError` for a `locale` that is not a locale tag.
    #[new]
    #[pyo3(signature = (source, *, locale=None, now=None, seed=None))]
    fn new(
        source: &Bound<'_, PyAny>,
        locale: Option<&str>,
        now: Option<f64>,
        seed: Option<u64>,
    ) -> PyResult<Self> {
        let locale = locale
            .map(|tag| {
                tag.parse::<sheetrs_io::locale::LocaleTag>()
                    .map_err(|error| PyValueError::new_err(error.to_string()))
            })
            .transpose()?;
        let with_locale = |book: Book| {
            let book = book.with_volatile(sheetrs_io::Volatile { now, seed });
            match &locale {
                Some(tag) => book.with_locale(tag.clone()),
                None => book,
            }
        };
        let mut book = if let Ok(bytes) = source.cast::<PyBytes>() {
            Self::from_source(with_locale(
                Book::from_bytes(bytes.as_bytes().to_vec()).map_err(format_error)?,
            ))
        } else if let Ok(bytes) = source.cast::<PyByteArray>() {
            Self::from_source(with_locale(
                Book::from_bytes(bytes.to_vec()).map_err(format_error)?,
            ))
        } else if source.hasattr("read")? {
            let data = read_stream(source)?;
            let mut book =
                Self::from_source(with_locale(Book::from_bytes(data).map_err(format_error)?));
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
            Self::from_source(with_locale(opened))
        };
        book.report = None;
        Ok(book)
    }

    /// The locale given to the workbook (`Book(..., locale=...)`), normalized as `"es-CO"`; `None` when
    /// it was not given one and has the one of its format.
    #[getter]
    fn locale(&self) -> Option<String> {
        self.source.locale().map(ToString::to_string)
    }

    /// The serial number of `NOW()` and `TODAY()` given to the workbook (`Book(..., now=...)`); `None`
    /// when it was not given one and the engine reads the clock.
    #[getter]
    fn now(&self) -> Option<f64> {
        self.source.volatile().now
    }

    /// The seed of `RAND()` given to the workbook (`Book(..., seed=...)`); `None` when it was not
    /// given one and the engine seeds from the clock.
    #[getter]
    fn seed(&self) -> Option<u64> {
        self.source.volatile().seed
    }

    /// `"xlsx"`, `"xlsm"` or `"ods"`.
    #[getter]
    fn format(&self) -> &'static str {
        self.source.format().extension()
    }

    /// The statistics of the workbook, as `sheetrs stats` reports them: the totals, the sheets, the
    /// macros and the defined names. It reads the workbook, in a single streaming pass, each time it is
    /// called; nothing else of the `Book` holds what it found. Raises `PendingOperationsError` with
    /// operations pending.
    fn stats(&self, py: Python<'_>) -> PyResult<PyStats> {
        self.queryable()?;
        let source = &self.source;
        let stats = py.detach(|| source.stats().map_err(to_py_err))?;
        Ok(PyStats::of(Arc::new(stats)))
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

    /// Hide rows (`"Sheet!3:10"`, or `"Sheet!3"`): *Format ▸ Rows ▸ Hide*. They stay in the file.
    /// One or an iterable of them, in the rows of the workbook as it is.
    fn hide_rows<'py>(
        slf: PyRefMut<'py, Self>,
        rows: &Bound<'_, PyAny>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        Self::rows_visibility(slf, rows, false)
    }

    /// Show rows, whatever hid them (by hand or by a filter): *Format ▸ Rows ▸ Show*.
    fn show_rows<'py>(
        slf: PyRefMut<'py, Self>,
        rows: &Bound<'_, PyAny>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        Self::rows_visibility(slf, rows, true)
    }

    /// Hide columns (`"Sheet!B:D"`, or `"Sheet!B"`): *Format ▸ Columns ▸ Hide*.
    fn hide_columns<'py>(
        slf: PyRefMut<'py, Self>,
        columns: &Bound<'_, PyAny>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        Self::columns_visibility(slf, columns, false)
    }

    /// Show columns: *Format ▸ Columns ▸ Show*.
    fn show_columns<'py>(
        slf: PyRefMut<'py, Self>,
        columns: &Bound<'_, PyAny>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        Self::columns_visibility(slf, columns, true)
    }

    /// Hide sheets by name (*Sheet ▸ Hide Sheet*): one name or an iterable of them. One sheet must
    /// stay visible.
    fn hide_sheets<'py>(
        slf: PyRefMut<'py, Self>,
        names: &Bound<'_, PyAny>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        Self::sheets_visibility(slf, names, false)
    }

    /// Show sheets by name (*Sheet ▸ Show Sheet*).
    fn show_sheets<'py>(
        slf: PyRefMut<'py, Self>,
        names: &Bound<'_, PyAny>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        Self::sheets_visibility(slf, names, true)
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
    /// or, with none declared, its formula cells. A range outside the declared outputs is answered
    /// from the workbook as it was written (the workbook is its own template here), calculated again
    /// with that range as the output. `major` is `"rows"` (the default) or `"columns"`: how each
    /// range is listed, as the `majorDimension` of the request.
    ///
    /// `now` (the serial number that `NOW()` and `TODAY()` give) and `seed` (the seed of `RAND()`)
    /// fix the volatile functions, as `--now` and `--seed` of `sheetc` do; without them the engine
    /// reads the clock. They also fix the recalculation that the same `save` runs after an edit that
    /// leaves results stale.
    ///
    /// Where the engine cannot calculate the workbook (a function it does not have, a reference it
    /// does not read) `save` still writes it, marked to be recalculated when it is opened, and issues
    /// a `RecalculationWarning`: the recalculation that is needed will be done by the spreadsheet
    /// program that opens the file later. With `to_dict()` or `save(json=...)` there is no answer to
    /// give, and it is an error.
    #[pyo3(signature = (inputs=None, ranges=None, *, major="rows", now=None, seed=None))]
    fn calculate<'py>(
        mut slf: PyRefMut<'py, Self>,
        py: Python<'py>,
        inputs: Option<&Bound<'py, PyDict>>,
        ranges: Option<&Bound<'py, PyAny>>,
        major: &str,
        now: Option<f64>,
        seed: Option<u64>,
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
            now,
            seed,
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
            let mut saved = Self::from_source(
                Book::from_bytes(bytes)
                    .map_err(to_py_err)?
                    .with_locale_of(&source)
                    .with_volatile_of(&source),
            );
            saved.stream = Some(stream.clone().unbind());
            saved.report = Some(report);
            Bound::new(py, saved)?.into_any()
        } else if let Some(path) = path_dest {
            let saved_source = Book::from_file(path)
                .map_err(to_py_err)?
                .with_locale_of(&source)
                .with_volatile_of(&source);
            let mut saved = Self::from_source(saved_source);
            saved.report = Some(report);
            Bound::new(py, saved)?.into_any()
        } else {
            if let Some(stream) = &source_stream {
                let bytes = done
                    .bytes
                    .ok_or_else(|| raise("Error", "the workbook was not written"))?;
                rewrite_stream(py, stream, &bytes)?;
                slf.source = Book::from_bytes(bytes)
                    .map_err(to_py_err)?
                    .with_locale_of(&source)
                    .with_volatile_of(&source);
            }
            slf.steps.clear();
            slf.report = Some(report);
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
/// default), and `now` and `seed` for the volatile functions.
pub(crate) fn evaluate(
    py: Python<'_>,
    formula: &str,
    locale: Option<&str>,
    now: Option<f64>,
    seed: Option<u64>,
) -> PyResult<Py<PyAny>> {
    let mut options = Options::default();
    options.config.now = now;
    options.config.seed = seed;
    let value = py
        .detach(|| memory::evaluate(formula, locale, &options))
        .map_err(to_py_err)?;
    json_to_py(py, &value)
}

//! What a query gives back: `CellError`, `Block` (the cells of one range) and `Match` (one cell a
//! search found).

use pyo3::exceptions::{PyAttributeError, PyIndexError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PySlice};
use std::hash::{DefaultHasher, Hash, Hasher};

/// A cell that holds an error: `#DIV/0!`, `#N/A`, or the code only `LibreOffice` has (`Err:502`).
///
/// It is a type of its own so that an error is not taken for a text that starts with `#`: a cell
/// with the text `#N/A` is a `str`, and one that calculated to `#N/A` is a `CellError`. Two errors
/// are equal when their codes are; an error is never equal to a `str`.
#[pyclass(name = "CellError", frozen, module = "sheetrs")]
pub struct CellError {
    code: String,
}

impl CellError {
    pub(crate) fn new(code: String) -> Self {
        Self { code }
    }
}

#[pymethods]
impl CellError {
    #[new]
    fn py_new(code: String) -> Self {
        Self { code }
    }

    /// The code the application shows: `#DIV/0!`, `#N/A`, `#NAME?`, `Err:502`…
    #[getter]
    fn code(&self) -> &str {
        &self.code
    }

    fn __str__(&self) -> &str {
        &self.code
    }

    fn __repr__(&self) -> String {
        format!(
            "CellError('{}')",
            self.code.replace('\\', "\\\\").replace('\'', "\\'")
        )
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other
            .cast::<Self>()
            .is_ok_and(|other| other.get().code == self.code)
    }

    fn __hash__(&self) -> u64 {
        let mut hasher = DefaultHasher::new();
        self.code.hash(&mut hasher);
        hasher.finish()
    }
}

/// The properties of a block, as the lists of rows they are.
pub(crate) struct Grids {
    pub values: Option<Py<PyList>>,
    pub formulas: Option<Py<PyList>>,
    pub number_formats: Option<Py<PyList>>,
    pub displays: Option<Py<PyList>>,
    pub comments: Option<Py<PyList>>,
}

/// The cells of one range: one list of rows per property that was asked for.
///
/// A bounded range (`A1:C10`) is as big as the range, with `None` where a cell is empty; a range
/// open at one end (`A:C`, `3:10`) is cut where its last non-empty cell is. With `major="columns"`
/// the lists are lists of columns. It iterates and indexes the rows of `values` (of the first
/// property asked for when `value` was not), has a `len`, and `to_dicts()` takes the first row as
/// the header. A property that was not asked for raises `AttributeError`.
#[pyclass(name = "Block", frozen, module = "sheetrs")]
pub struct Block {
    range: String,
    major: &'static str,
    grids: Grids,
}

impl Block {
    pub(crate) fn new(range: String, major: &'static str, grids: Grids) -> Self {
        Self {
            range,
            major,
            grids,
        }
    }

    fn primary<'py>(&self, py: Python<'py>) -> Bound<'py, PyList> {
        let g = &self.grids;
        [
            &g.values,
            &g.formulas,
            &g.number_formats,
            &g.displays,
            &g.comments,
        ]
        .into_iter()
        .flatten()
        .next()
        .map_or_else(|| PyList::empty(py), |list| list.bind(py).clone())
    }

    fn property<'py>(
        py: Python<'py>,
        grid: Option<&Py<PyList>>,
        name: &str,
        asked: &str,
    ) -> PyResult<Bound<'py, PyList>> {
        grid.map(|list| list.bind(py).clone()).ok_or_else(|| {
            PyAttributeError::new_err(format!(
                "`{name}` was not asked for: add '{asked}' to `properties`"
            ))
        })
    }
}

#[pymethods]
impl Block {
    /// The range this block holds, as it was asked for.
    #[getter]
    fn range(&self) -> &str {
        &self.range
    }

    /// `"rows"` or `"columns"`: what the lists of the properties are lists of.
    #[getter]
    fn major(&self) -> &'static str {
        self.major
    }

    /// The values: lists of rows; `None` for an empty cell.
    #[getter]
    fn values<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        Self::property(py, self.grids.values.as_ref(), "values", "value")
    }

    /// The formulas, without the `=`, in the syntax of the file; `None` for a cell with none.
    #[getter]
    fn formulas<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        Self::property(py, self.grids.formulas.as_ref(), "formulas", "formula")
    }

    #[getter]
    fn number_formats<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        Self::property(
            py,
            self.grids.number_formats.as_ref(),
            "number_formats",
            "number_format",
        )
    }

    /// What a person sees in each cell.
    #[getter]
    fn displays<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        Self::property(py, self.grids.displays.as_ref(), "displays", "display")
    }

    #[getter]
    fn comments<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        Self::property(py, self.grids.comments.as_ref(), "comments", "comment")
    }

    fn __len__(&self, py: Python<'_>) -> usize {
        self.primary(py).len()
    }

    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        self.primary(py).try_iter().map(Bound::into_any)
    }

    fn __getitem__<'py>(
        &self,
        py: Python<'py>,
        index: &Bound<'py, PyAny>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let rows = self.primary(py);
        if index.is_instance_of::<PySlice>() {
            return rows.as_any().get_item(index);
        }
        rows.as_any()
            .get_item(index)
            .map_err(|_| PyIndexError::new_err("block index out of range"))
    }

    /// The rows as dicts, the first row being the header: `[{"Name": "a", "Qty": 1}, ...]`. The keys
    /// are the header cells as `str`. Needs `value` and `major="rows"`.
    fn to_dicts<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        if self.major != "rows" {
            return Err(PyAttributeError::new_err(
                "`to_dicts` needs `major=\"rows\"`",
            ));
        }
        let Some(rows) = self.grids.values.as_ref().map(|list| list.bind(py).clone()) else {
            return Err(PyAttributeError::new_err(
                "`to_dicts` needs 'value' in `properties`",
            ));
        };
        let out = PyList::empty(py);
        let mut header: Option<Vec<Bound<'py, PyAny>>> = None;
        for row in rows.iter() {
            let cells: Vec<Bound<'py, PyAny>> = row.try_iter()?.collect::<PyResult<_>>()?;
            match &header {
                None => {
                    header = Some(
                        cells
                            .iter()
                            .map(|c| c.str().map(Bound::into_any))
                            .collect::<PyResult<_>>()?,
                    );
                }
                Some(keys) => {
                    let item = PyDict::new(py);
                    for (key, cell) in keys.iter().zip(cells) {
                        item.set_item(key, cell)?;
                    }
                    out.append(item)?;
                }
            }
        }
        Ok(out)
    }

    fn __repr__(&self, py: Python<'_>) -> String {
        format!(
            "Block({:?}, {} {})",
            self.range,
            self.primary(py).len(),
            self.major
        )
    }
}

/// One cell a search found, with the properties that were asked for (`None` for the others).
#[pyclass(name = "Match", frozen, module = "sheetrs")]
pub struct Match {
    /// The sheet, shared by the matches of it.
    pub(crate) sheet: std::sync::Arc<str>,
    /// One-based.
    #[pyo3(get)]
    pub(crate) row: u32,
    /// One-based.
    #[pyo3(get)]
    pub(crate) col: u32,
    #[pyo3(get)]
    pub(crate) value: Py<PyAny>,
    #[pyo3(get)]
    pub(crate) formula: Option<String>,
    #[pyo3(get)]
    pub(crate) number_format: Option<String>,
    #[pyo3(get)]
    pub(crate) display: Option<String>,
    #[pyo3(get)]
    pub(crate) comment: Option<String>,
}

#[pymethods]
impl Match {
    /// The sheet.
    #[getter]
    fn sheet(&self) -> &str {
        &self.sheet
    }

    /// The cell in A1 notation, worked out from `row` and `col`.
    #[getter]
    fn cell(&self) -> String {
        sheetrs_formula::addr::a1(self.row.saturating_sub(1), self.col.saturating_sub(1))
    }

    fn __repr__(&self, py: Python<'_>) -> String {
        let value = self
            .value
            .bind(py)
            .repr()
            .map_or_else(|_| "?".to_owned(), |r| r.to_string());
        format!("Match({}!{}, {value})", self.sheet, self.cell())
    }
}

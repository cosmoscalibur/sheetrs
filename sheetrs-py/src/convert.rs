//! JSON to Python: the values of a cell as the Python types the package promises.
//!
//! A number is a `float`, or an `int` when it is an exact integer; text is `str`, a boolean `bool`,
//! an empty cell `None`, an error a [`CellError`], and a date, a time or both (when
//! the call asked for `dates=True`) a `datetime.date`, `datetime.time` or `datetime.datetime`.

use crate::CellError;
use pyo3::exceptions::{PyOverflowError, PyTypeError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDict, PyList, PyString, PyTuple};
use serde_json::{Map, Number, Value};

/// A number as Python reads it: an `int` when it is an exact integer a double holds, a `float`
/// otherwise.
fn number(py: Python<'_>, n: &Number) -> PyResult<Py<PyAny>> {
    if let Some(i) = n.as_i64() {
        return Ok(i.into_pyobject(py)?.into_any().unbind());
    }
    if let Some(u) = n.as_u64() {
        return Ok(u.into_pyobject(py)?.into_any().unbind());
    }
    let x = n.as_f64().unwrap_or(f64::NAN);
    if x.fract() == 0.0 && x.abs() < 1e15 {
        // `as` is exact here: an integer below 1e15 fits an i64.
        #[allow(clippy::cast_possible_truncation)]
        return Ok((x as i64).into_pyobject(py)?.into_any().unbind());
    }
    Ok(x.into_pyobject(py)?.into_any().unbind())
}

/// The object a JSON object stands for: an error, or a date or time with its type.
fn object(py: Python<'_>, o: &Map<String, Value>) -> PyResult<Py<PyAny>> {
    if let Some(Value::String(error)) = o.get("error") {
        // An error only `LibreOffice` has carries its code next to the name Excel would give it.
        let code = match o.get("code") {
            Some(Value::String(code)) => code.clone(),
            _ => error.clone(),
        };
        return Ok(Py::new(py, CellError::new(code))?.into_any());
    }
    if let (Some(Value::String(kind)), Some(Value::String(text))) = (o.get("type"), o.get("value"))
    {
        let datetime = py.import("datetime")?;
        let class = match kind.as_str() {
            "date" => "date",
            "time" => "time",
            "datetime" => "datetime",
            _ => {
                return Err(PyTypeError::new_err(format!(
                    "unknown typed value `{kind}`"
                )));
            }
        };
        return Ok(datetime
            .getattr(class)?
            .call_method1("fromisoformat", (text,))?
            .unbind());
    }
    let dict = PyDict::new(py);
    for (key, value) in o {
        dict.set_item(key, json_to_py(py, value)?)?;
    }
    Ok(dict.into_any().unbind())
}

/// The Python value of a JSON value of the tools (a cell, or what a program answers).
pub(crate) fn json_to_py(py: Python<'_>, value: &Value) -> PyResult<Py<PyAny>> {
    Ok(match value {
        Value::Null => py.None(),
        Value::Bool(b) => b.into_pyobject(py)?.to_owned().into_any().unbind(),
        Value::Number(n) => number(py, n)?,
        Value::String(s) => s.as_str().into_pyobject(py)?.into_any().unbind(),
        Value::Array(items) => {
            let list = PyList::empty(py);
            for item in items {
                list.append(json_to_py(py, item)?)?;
            }
            list.into_any().unbind()
        }
        Value::Object(o) => object(py, o)?,
    })
}

/// The JSON of a value a caller gives a program as an input: `None`, `bool`, `int`, `float`, `str`,
/// and lists or tuples of them (an array of rows, for an input that is a range).
pub(crate) fn py_to_json(value: &Bound<'_, PyAny>) -> PyResult<Value> {
    if value.is_none() {
        return Ok(Value::Null);
    }
    if let Ok(b) = value.cast::<PyBool>() {
        return Ok(Value::Bool(b.is_true()));
    }
    if value.is_instance_of::<pyo3::types::PyInt>() {
        let i: i64 = value
            .extract()
            .map_err(|_| PyOverflowError::new_err("an integer input does not fit in 64 bits"))?;
        return Ok(Value::from(i));
    }
    if value.is_instance_of::<pyo3::types::PyFloat>() {
        let x: f64 = value.extract()?;
        return Number::from_f64(x)
            .map(Value::Number)
            .ok_or_else(|| PyTypeError::new_err("an input must be a finite number"));
    }
    if let Ok(text) = value.cast::<PyString>() {
        return Ok(Value::String(text.to_str()?.to_owned()));
    }
    if value.is_instance_of::<PyList>() || value.is_instance_of::<PyTuple>() {
        return value
            .try_iter()?
            .map(|item| py_to_json(&item?))
            .collect::<PyResult<Vec<_>>>()
            .map(Value::Array);
    }
    Err(PyTypeError::new_err(format!(
        "an input is None, a bool, an int, a float, a str, or a list of them, not {}",
        value.get_type().name()?
    )))
}

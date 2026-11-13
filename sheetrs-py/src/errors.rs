//! The exceptions: the hierarchy is defined in `sheetrs/_errors.py` (so that a class can be an
//! `Error` and a `ValueError` at once) and raised from here.

use pyo3::exceptions::PyRuntimeError;
use pyo3::prelude::*;

/// An exception of the package by its name (`"Error"`, `"FormatError"`, `"InvalidRangeError"`,
/// `"UnsupportedDateSystemError"`, `"UnavailableFunctionError"`, `"PendingOperationsError"`) with this
/// message.
pub(crate) fn raise(name: &str, message: impl Into<String>) -> PyErr {
    let message = message.into();
    Python::attach(|py| {
        py.import("sheetrs._errors")
            .and_then(|module| module.getattr(name))
            .and_then(|class| class.call1((message.clone(),)))
            .map_or_else(|_| PyRuntimeError::new_err(message), PyErr::from_value)
    })
}

/// A failure of the file system anywhere in the chain of causes becomes `OSError`
/// (`FileNotFoundError`, …), keeping its kind; anything else, `sheetrs.Error` (or its subclasses
/// `UnsupportedDateSystemError` when the message says the workbook has a date system that is not
/// supported, and `UnavailableFunctionError` when it says the engine lacks a function). The message
/// is the whole chain.
pub(crate) fn to_py_err<E: std::error::Error + 'static>(error: E) -> PyErr {
    let mut cause: Option<&(dyn std::error::Error + 'static)> = Some(&error);
    while let Some(current) = cause {
        if let Some(io) = current.downcast_ref::<std::io::Error>() {
            return PyErr::from(std::io::Error::new(io.kind(), format!("{error:#}")));
        }
        cause = current.source();
    }
    classified(format!("{error:#}"))
}

/// The exception a message of the tools stands for (see [`to_py_err`]): for a failure that was kept as
/// text.
pub(crate) fn classified(message: String) -> PyErr {
    if is_unsupported_date_system(&message) {
        raise("UnsupportedDateSystemError", message)
    } else if is_unavailable_function(&message) {
        raise("UnavailableFunctionError", message)
    } else {
        raise("Error", message)
    }
}

/// Whether a message says that the workbook numbers its dates from a date that is not one of the three
/// date systems of `LibreOffice` Calc (the message of `sheetrs_io::UnsupportedDateSystem`, which names
/// the date).
fn is_unsupported_date_system(message: &str) -> bool {
    message.contains("unsupported date system:")
}

/// Whether a message of the compiler says that the workbook calls functions of `LibreOffice` that the
/// engine does not implement.
fn is_unavailable_function(message: &str) -> bool {
    message.contains("that this engine does not implement")
}

pub(crate) fn invalid_range(message: impl Into<String>) -> PyErr {
    raise("InvalidRangeError", message)
}

pub(crate) fn format_error(message: impl std::fmt::Display) -> PyErr {
    raise("FormatError", message.to_string())
}

pub(crate) fn pending(operations: usize) -> PyErr {
    let plural = if operations == 1 { "" } else { "s" };
    raise(
        "PendingOperationsError",
        format!(
            "the Book has {operations} pending operation{plural} that nothing has run yet. If the \
             query depends on them, call save() first and query the Book it returns; if it does \
             not, run the query before adding the operations"
        ),
    )
}

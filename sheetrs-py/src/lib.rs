//! Python bindings for the `SheetRS` suite.
//!
//! The extension module is `sheetrs._sheetrs`; the `sheetrs` package re-exports it with the
//! exceptions of `sheetrs/_errors.py`. The API is one object, `Book` (`book::PyBook`): a workbook (a
//! path, bytes or a binary stream) that reads nothing when it is made and keeps nothing of the workbook,
//! whose
//! operations accumulate until `save`, the cut that plans and runs them, and whose queries (`range`,
//! `find`, `lint`, `inspect`, `stats`) read what they need; `calculate` is one of the operations, and `value_ranges()`
//! or `save(json=…)` give its value ranges. The results are typed values (`results`: `Stats`, `Sheet`,
//! `ValueRanges`, `Report`, `Diagnostic`, `Inspection`; `lint`: `LintReport`) that write themselves with
//! `to_dict()` and `to_json()`. The logic is in the Rust crates (`sheetrs-io`, `sheetrs-query`,
//! `sheetrs-edit`, `sheetrs-compiler`, `sheetrs-engine`, `sheetrs-linter`, `sheetrs-stats`); this crate
//! converts between them and Python.
//!
//! The functions release the GIL while they read, analyze, calculate and write a workbook.

mod book;
mod cells;
mod convert;
mod errors;
mod lint;
mod plan;
mod results;

use cells::{Block, CellError, Match};
use pyo3::prelude::*;

/// The value of one formula, as `sheetc evaluate` gives it, with the regional settings of `locale`
/// (a BCP-47 tag, `"en"` by default): `evaluate("=SUM(1,2)")` is `3`. An error is a `CellError`.
/// `now` (the serial number of `NOW()` and `TODAY()`) and `seed` (the seed of `RAND()`) fix the
/// volatile functions; without them the clock decides.
#[pyfunction]
#[pyo3(name = "evaluate", signature = (formula, *, locale=None, now=None, seed=None))]
fn evaluate_formula(
    py: Python<'_>,
    formula: &str,
    locale: Option<&str>,
    now: Option<f64>,
    seed: Option<u64>,
) -> PyResult<Py<PyAny>> {
    book::evaluate(py, formula, locale, now, seed)
}

/// The rules of the linter, ordered by id: a list of dicts with `id`, `name`, `category`,
/// `default` (runs without configuration) and `enabled` (runs with the configuration of the file
/// `config_path`). The ids and categories are what `disabled_rules` and `enabled_rules` accept.
#[pyfunction]
#[pyo3(signature = (config_path=None))]
#[allow(clippy::needless_pass_by_value)] // pyo3 hands over an owned argument
fn lint_rules<'py>(
    py: Python<'py>,
    config_path: Option<std::borrow::Cow<'_, std::path::Path>>,
) -> PyResult<Bound<'py, pyo3::types::PyList>> {
    lint::rules(py, config_path.as_deref())
}

#[pymodule(gil_used = false)]
mod _sheetrs {
    use pyo3::prelude::*;

    #[pymodule_export]
    use super::book::PyBook;
    #[pymodule_export]
    use super::lint::{PyLintReport, PyViolation};
    #[pymodule_export]
    use super::results::{
        PyDiagnostic, PyFunctionUse, PyInspection, PyReport, PySheet, PySheets, PySkipped, PyStats,
        PyValueRanges,
    };
    #[pymodule_export]
    use super::{Block, CellError, Match, evaluate_formula, lint_rules};

    #[pymodule_init]
    fn init(module: &Bound<'_, PyModule>) -> PyResult<()> {
        module.add("__version__", env!("CARGO_PKG_VERSION"))
    }
}

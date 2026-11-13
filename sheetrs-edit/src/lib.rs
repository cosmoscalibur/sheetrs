//! sheetrs-edit: modifications of XLSX and ODS workbooks.
//!
//! The library behind the `sheetrs edit` executable and the language bindings. [`BookEdit`] adds the
//! changes to a [`Book`](sheetrs_io::Book) (sheets and names deleted, formulas replaced by their
//! values, rows and columns deleted, text found and replaced as `LibreOffice` does it:
//! [`BookEdit::replace`]) and writes the result, streaming the archive without rewriting what it
//! does not change: to a new file ([`BookEdit::write`]), over the original
//! ([`BookEdit::write_in_place`]) or into memory ([`BookEdit::write_to`]).
//!
//! The changes are written against the workbook as it is, and the order they are given in changes
//! nothing: the planner (module `plan`) runs them in the fewest passes over the file, in an order of
//! its own ([`BookEdit::explain`] says which). To work step by step, save between steps.
//!
//! A conversion of formulas to values reads the results a workbook caches. After a change that
//! leaves them stale (rows deleted, a sheet removed) they are not the ones an application shows, so
//! the results are calculated again first, by the engine of `sheetrs-compiler` (the semantics of
//! `LibreOffice`); stale values are never written.
//!
//! [`BookEdit::calculate`] is an operation: the engine writes the inputs and calculates every
//! formula, and the workbook is saved with the results and without the mark to recalculate when it
//! is opened. Where the engine cannot calculate it, the mark is set and the run says so in a
//! [`Diagnostic`](sheetrs_protocol::Diagnostic) of the [`Report`] (code [`RECALC_NOT_SUPPORTED`]). The value ranges of a calculation are in the
//! [`Report`] (the `answer` of the [`Calculation`]) as [`ValueRanges`](sheetrs_protocol::ValueRanges), which writes them as JSON
//! ([`Report::write_json`]).

mod book_edit;
mod calculate;
#[cfg(feature = "cli")]
pub mod cli;
mod diagnostics;
mod error;
mod plan;
mod replace;
mod stale;

#[cfg(test)]
mod tests;

pub use book_edit::BookEdit;
pub use diagnostics::RECALC_NOT_SUPPORTED;
pub use error::{Error, Result, ScopeError};
pub use plan::explain;
pub use replace::{Report, Skipped};
pub use sheetrs_io::write::{CachedResult, Calculation, CellWrites, Constant, ValueScope};
pub use sheetrs_io::write::{ColumnRange, RowRange, VisibilityChange};

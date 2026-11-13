//! Reading the cells of a workbook, and finding the ones that match a search.
//!
//! [`BookQuery::query`] answers a [`Query`]: the scopes (a range, a sheet, the
//! workbook), the properties of each cell (`value`, `formula`, `number_format`, `display`) and,
//! optionally, a [`Find`] that keeps the cells that match. Only what is asked
//! is read: the readers skip the sheets that are not in scope, stop after the last row a window
//! reaches, and do not build the properties that are not asked for.
//!
//! Numbers are given as the serial the file stores; `dates` gives the ones whose format is a
//! date or a time as a [`DateValue`] (ISO 8601 text with its type), worked out with the date system of
//! the file.
//!
//! The result is the one `LibreOffice` gives for that file (an ODS keeps the dates as dates and its own
//! errors, `Err:502`; an XLSX has neither).
//!
//! [`text`] holds what a cell says as text (what a person sees, what a search sees) and
//! `matcher` the matching of a search against a text; both are shared with `sheetrs-edit`,
//! which replaces what this crate finds.

mod blocks;
#[cfg(feature = "cli")]
pub mod cli;
mod dates;
mod error;
mod inherit;
mod matcher;
mod run;
pub mod text;

pub use blocks::{Block, FindOptions, Major, RangeOptions};
pub use error::{Error, Result};
pub use matcher::Matcher;
pub use run::BookQuery;
pub use sheetrs_protocol::{CellRecord, DateValue, Find, Property, Query, QueryResponse, Scope};

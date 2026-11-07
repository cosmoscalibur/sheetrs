//! Reading the cells of a workbook, and finding the ones that match a search.
//!
//! [`run`] answers a [`Query`](sheetrs_protocol::Query): the scopes (a range, a sheet, the
//! workbook), the properties of each cell (`value`, `formula`, `number_format`, `display`) and,
//! optionally, a [`Find`](sheetrs_protocol::Find) that keeps the cells that match. Only what is asked
//! is read: the readers skip the sheets that are not in scope, stop after the last row a window
//! reaches, and do not build the properties that are not asked for.
//!
//! The result is the same for an XLSX and an ODS wherever the two files say the same thing; where
//! `LibreOffice` itself treats them differently (an ODS keeps the dates as dates and its own errors,
//! `Err:502`) the result follows `LibreOffice`.
//!
//! [`text`] holds what a cell says as text (what a person sees, what a search sees) and
//! [`matcher`] the matching of a search against a text; both are shared with `sheetrs-edit`,
//! which replaces what this crate finds.

mod error;
pub mod matcher;
mod run;
pub mod text;

pub use error::{Error, Result};
pub use matcher::Matcher;
pub use run::run;
pub use sheetrs_protocol::{CellRecord, Find, Property, Query, QueryResponse, Scope};

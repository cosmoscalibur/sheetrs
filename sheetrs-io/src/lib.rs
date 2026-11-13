//! sheetrs-io: the input of the `SheetRS` suite (named `sheetrs-core` before 0.4.0), and its shared
//! kernel.
//!
//! It holds what every tool needs and no tool owns: the format-agnostic workbook model
//! ([`reader::Workbook`], [`reader::Cell`], [`reader::Grid`]) and the streaming XLSX/ODS readers
//! that fill it. Everything a tool decides with the model lives in that tool's crate: the rules
//! and violations in `sheetrs-linter`, the metrics in `sheetrs-stats`, the modifications in
//! `sheetrs-edit`, the compilation in `sheetrs-compiler`.
//!
//! The feature `write` adds the other half: the writers of XLSX and ODS ([`mod@write`]), which rewrite a
//! workbook as a stream. It is off by default: the tools that only read do not carry it.

mod book;
mod entry_format;
mod error;
pub mod reader;
#[cfg(feature = "write")]
pub mod write;

pub use book::Book;
pub use entry_format::EntryFormat;
pub use error::{Error, Result};
pub use sheetrs_formula::{DateSystem, UnsupportedDateSystem};

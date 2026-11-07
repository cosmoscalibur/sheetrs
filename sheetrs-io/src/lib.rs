//! sheetrs-io: the input of the `SheetRS` suite (named `sheetrs-core` before 0.4.0), and its shared
//! kernel.
//!
//! It holds what every tool needs and no tool owns: the format-agnostic workbook model
//! ([`reader::Workbook`], [`reader::Cell`], [`reader::Grid`]) and the streaming XLSX/ODS readers
//! that fill it. Everything a tool decides with the model lives in that tool's crate: the rules
//! and violations in `sheetrs-linter`, the metrics in `sheetrs-stats`, the modifications in
//! `sheetrs-edit`, the compilation in `sheetrs-compiler`.

mod error;
pub mod reader;

pub use error::{Error, Result};

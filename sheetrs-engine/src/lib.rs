//! The calculation engine of compiled spreadsheets.
//!
//! A [`Program`] (the format-independent IR produced by `sheetrs-compiler`) is
//! executed by a demand-driven [`engine::Engine`] with JSON in and JSON out. This
//! crate has no knowledge of XLSX or ODS: it holds the IR, the function
//! registry (the namespaces of `OpenDocument`: `core`, `microsoft`, `openoffice`, `libreoffice`, `legacy`), values, dates, text conversion,
//! collation, search patterns and the run-time command line used by compiled
//! executables (`sheetrs-rt`).
//!
//! See `sheetrs-compiler/docs/DESIGN.md` and `FIDELITY.md`.

pub mod addr;
pub mod bundle;
pub mod calculate;
mod casemap;
pub mod cli;
pub mod clock;
pub mod coerce;
mod collate;
pub mod consts;
pub mod dates;
pub mod datetext;
pub mod engine;
mod error;
pub mod fastmap;
pub mod format;
pub mod functions;
pub mod ir;
mod locale_data;
pub mod matching;
mod notes;
pub mod options;
pub mod regex_subset;
pub mod value;

pub use calculate::Overrides;
pub use error::{Error, Result};
pub use ir::Program;
pub use value::Value;

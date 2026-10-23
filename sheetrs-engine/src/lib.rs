//! The calculation engine of compiled spreadsheets.
//!
//! A [`Program`] (the format-independent IR produced by `sheetrs-compiler`) is
//! executed by a demand-driven [`engine::Engine`] with JSON in and JSON out. This
//! crate has no knowledge of XLSX or ODS: it holds the IR, the function
//! registry (`core`/`xlsx`/`ods` namespaces), values, dates, text conversion,
//! collation, search patterns and the run-time command line used by compiled
//! executables (`sheetrs-rt`).
//!
//! See `sheetrs-compiler/docs/DESIGN.md` and `FIDELITY.md`.

pub mod addr;
pub mod bundle;
pub mod cli;
pub mod coerce;
pub mod dates;
pub mod engine;
pub mod fastmap;
pub mod format;
pub mod functions;
pub mod ir;
pub mod matching;
pub mod options;
pub mod run;
pub mod value;

pub use ir::Program;
pub use run::Overrides;
pub use value::Value;

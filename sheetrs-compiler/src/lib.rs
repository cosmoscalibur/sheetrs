//! sheetrs-compiler: compile spreadsheets (XLSX/ODS) into standalone
//! executables.
//!
//! ```text
//! file ──loader──▶ Model ──compile──▶ Program (IR) ──bundle──▶ executable
//!                    │                    │
//!                 (format-specific)   (format-independent; run by sheetrs-engine)
//! ```
//!
//! The IR, the engine and the function library live in `sheetrs-engine`; this
//! crate reads workbooks, lowers formulas (parsed by `sheetrs-formula`), orders
//! and prunes them. See `docs/DESIGN.md`.

pub mod cellstore;
pub mod compile;
pub mod config;
mod error;
pub mod formula;
pub mod graph;
#[cfg(feature = "load")]
pub mod loader;
pub mod locale;
pub mod model;
pub mod recalc;
pub mod settings;
pub mod verify;

pub use compile::{BindingSpec, Features, Report, Spec, compile};
pub use error::{Error, Result};
pub use model::Model;

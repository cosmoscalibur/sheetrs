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

mod cellstore;
pub mod compile;
pub mod config;
mod error;
mod formula;
mod graph;
#[cfg(feature = "load")]
pub mod loader;
pub mod locale;
#[cfg(feature = "load")]
pub mod memory;
pub mod model;
pub mod recalc;
pub mod settings;
pub mod verify;
pub mod wasm_target;

pub use compile::{BindingSpec, Features, ModelArg, Report, Spec, compile};
pub use error::{Error, Result};
pub use model::Model;

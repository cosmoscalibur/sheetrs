//! The contract of the `SheetRS` tools.
//!
//! What a caller sends and what it gets back, as typed JSON that every front end (the command line,
//! the Rust libraries, WebAssembly, Python) shares:
//!
//! - [`calculate`]: the request a compiled program answers (`inputs`, `ranges`, `majorDimension`) and its
//!   response (`valueRanges` or `outputs`).
//! - [`query`]: reading the cells of a workbook, by range, sheet or workbook, with the properties
//!   asked for (`value`, `formula`, `number_format`, `display`), or finding the cells that match.
//! - [`find`]: what is searched for and how, as `LibreOffice`'s Find & Replace says it; the same
//!   description serves a read-only search (`sheetrs-query`) and a replacement (`sheetrs-edit`).
//! - [`Diagnostic`]: what a run reports besides its result (level, stable code, message, context fields), and
//!   the [`Logger`] that the command lines write them with.
//!
//! The crate holds no behaviour but the writing of a diagnostic: no file is read, no formula is evaluated, no text is matched. It
//! depends on `sheetrs-formula` (the grammar of references) and on `serde`, so a client can speak
//! the protocol without any of the tools.
//!
//! ```
//! use sheetrs_protocol::{Property, Query, Scope};
//!
//! // The request a client sends, as JSON, and the typed value both ends share.
//! let query: Query = serde_json::from_str(
//!     r#"{"scope": {"range": "Data!A1:C9"}, "properties": ["value", "formula"], "dates": true}"#,
//! )?;
//! assert_eq!(query.properties, [Property::Value, Property::Formula]);
//! assert!(query.dates);
//! assert!(matches!(&query.scope[0], Scope::Range(range) if range.sheet.as_deref() == Some("Data")));
//! # Ok::<(), serde_json::Error>(())
//! ```

pub mod calculate;
mod diagnostic;
mod error;
pub mod find;
pub mod query;
pub mod reference;

pub use calculate::{MajorDimension, Request, ValueRange, ValueRanges};
#[cfg(feature = "cli")]
pub use diagnostic::LogArgs;
pub use diagnostic::{Diagnostic, Level, LogFormat, Logger};
pub use error::{Error, Result};
pub use find::{Find, Mode, Replace, SearchIn};
pub use query::{CellRecord, DateValue, Property, Query, QueryResponse, Scope};
pub use reference::Reference;

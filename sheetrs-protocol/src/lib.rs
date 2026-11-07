//! The contract of the `SheetRS` tools.
//!
//! What a caller sends and what it gets back, as typed JSON that every front end (the command line,
//! the Rust libraries, WebAssembly, Python) shares:
//!
//! - [`run`]: the request a compiled program answers (`inputs`, `ranges`, `majorDimension`) and its
//!   response (`valueRanges` or `outputs`).
//! - [`query`]: reading the cells of a workbook, by range, sheet or workbook, with the properties
//!   asked for (`value`, `formula`, `number_format`, `display`), or finding the cells that match.
//! - [`find`]: what is searched for and how, as `LibreOffice`'s Find & Replace says it; the same
//!   description serves a read-only search (`sheetrs-query`) and a replacement (`sheetrs-edit`).
//!
//! The crate holds no behaviour: no file is read, no formula is evaluated, no text is matched. It
//! depends on `sheetrs-formula` (the grammar of references) and on `serde`, so a client can speak
//! the protocol without any of the tools.

mod error;
pub mod find;
pub mod query;
pub mod reference;
pub mod run;

pub use error::{Error, Result};
pub use find::{Find, Mode, Replace, SearchIn};
pub use query::{CellRecord, Property, Query, QueryResponse, Scope};
pub use reference::Reference;
pub use run::{MajorDimension, Request, Response, ValueRange};

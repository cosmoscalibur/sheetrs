//! sheetrs-edit: modifications of XLSX and ODS workbooks.
//!
//! The library behind the `sheetedit` executable and the language bindings. An [`Editor`] is
//! bound to a [`Source`](sheetrs_io::reader::Source), collects the changes to make (sheets and
//! names removed, formulas replaced by their values, rows and columns deleted, text found and
//! replaced as `LibreOffice` does it: [`Editor::replace`]) and saves the
//! result as a new file, streaming the archive without rewriting what it does not change.
//! The sequence of changes is planned into the fewest passes over the file that give the result
//! of running them one at a time; [`Editor::explain`] says what it will do.

mod editor;
mod error;
mod plan;
mod replace;
mod template;
mod writer;
mod writes;

#[cfg(test)]
mod tests;

pub use editor::Editor;
pub use error::{Error, Result, ScopeError};
pub use replace::{Report, Skipped};
pub use template::fill_template;
pub use writer::{CachedResult, ColumnRange, RowRange, ValueScope};
pub use writes::{CellWrites, Constant};

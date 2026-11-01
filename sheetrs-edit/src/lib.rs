//! sheetrs-edit: modifications of XLSX and ODS workbooks.
//!
//! The library behind the `sheetedit` executable and the language bindings. [`writer`] streams
//! the archive of a workbook into a modified copy (sheets and names removed, formulas replaced by
//! their values, rows and columns deleted), without rewriting what it does not change.
//! [`EditPlan`] collects the operations to run on one workbook and applies them in a single pass.

mod plan;
pub mod writer;

pub use plan::{EditPlan, Operation};
pub use writer::{ColumnRange, RowRange, ValueScope};

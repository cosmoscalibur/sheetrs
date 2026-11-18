//! Core types of sheetrs.

mod coord;
mod date;
mod store;
mod workbook;

pub use coord::{AxisRef, CellPos, CellRef, MAX_COLS, MAX_ROWS, RefError};
pub use store::{
    AbsorbError, CellView, FormulaRef, LimitError, MAX_ARRAY_CELLS, RowMajor, SharedMember,
    SheetBuilder, SheetView, ValueRef,
};
pub use workbook::{
    Cell, DateSystem, ErrorKind, ExternalLink, Formula, MacroKind, MacroModule, Value, Visibility,
    Workbook,
};

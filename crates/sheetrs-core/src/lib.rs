//! Core types of sheetrs.

mod coord;
mod date;
mod opaque;
mod store;
mod workbook;

pub use coord::{AxisRef, CellPos, CellRef, MAX_COLS, MAX_ROWS, RefError};
pub use opaque::{Frame, OpaquePackage, OpaquePart, PackageFormat, SheetOpaque, StyledCells};
pub use store::{
    AbsorbError, CellView, FormulaRef, LimitError, MAX_ARRAY_CELLS, RowMajor, SharedMember,
    SheetBuilder, SheetView, ValueRef,
};
pub use workbook::{
    Cell, DateSystem, ErrorKind, ExternalLink, Formula, FormulaGrammar, MacroKind, MacroModule,
    NamedExpression, Value, Visibility, Workbook,
};

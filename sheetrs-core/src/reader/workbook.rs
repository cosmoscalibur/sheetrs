//! Workbook data structures

use std::collections::HashMap;

use sheetrs_formula::grid::Grid;
use std::path::PathBuf;
use std::sync::Arc;

use chrono::{DateTime, Utc};

use crate::violation::ExcelError;

/// Represents an external workbook reference
#[derive(Debug, Clone)]
pub struct ExternalWorkbook {
    /// 0-based index in the collection (maps to XLSX 1-based index - 1)
    /// For XLSX: index N corresponds to [N+1] in formulas
    /// For ODS: index N corresponds to order of appearance in metadata
    pub index: usize,
    /// Basename of the external workbook (e.g., "Book1.xlsx")
    pub path: String,
}

#[derive(Debug, Clone, Default)]
/// The main structure representing a parsed spreadsheet workbook
pub struct Workbook {
    /// Absolute path to the workbook file
    pub path: PathBuf,
    /// Collection of worksheets in the workbook
    pub sheets: Vec<Sheet>,
    /// Map of defined names (named ranges) to their values/formulas
    pub defined_names: HashMap<String, String>,
    /// List of hidden sheet names
    pub hidden_sheets: Vec<String>,
    /// Whether the workbook contains macros or VBA code
    pub has_macros: bool,
    /// External workbooks referenced by this workbook (0-indexed collection)
    /// For XLSX: index N corresponds to [N+1] in formulas
    /// For ODS: index N corresponds to order of appearance in metadata
    pub external_workbooks: Vec<ExternalWorkbook>,
    /// Last modification date (parsed from file metadata)
    pub modified_date: Option<DateTime<Utc>>,
    /// Whether the workbook uses the 1904 date system (legacy Mac compatibility)
    pub date1904: bool,
    /// File size in bytes (extracted from file metadata)
    pub file_size_bytes: u64,
    /// Calculation and locale settings stored in the document
    pub settings: DocumentSettings,
    /// The document properties the file declares, as `(name, value)` with the values as the file
    /// writes them (an XLSX date ends in `Z`, an ODS one has no zone) in the order of the file, the
    /// keywords of an ODS, one element each, last (see
    /// [`WorkbookReader::read_properties`](super::WorkbookReader::read_properties)).
    pub properties: Vec<(String, String)>,
}

/// Iterative calculation settings (circular references).
#[derive(Debug, Clone, PartialEq, Default)]
pub struct IterationSettings {
    /// Whether iteration is switched on.
    pub enabled: bool,
    /// Maximum number of iterations (Excel and LibreOffice default: 100).
    pub max_iterations: u32,
    /// Maximum change between iterations (default: 0.001).
    pub max_change: f64,
}

/// Calculation settings a document carries, as stored. Every field is `None`
/// when the file does not state it: the reader applies no defaults, so the
/// consumer can tell "not stated" from "stated as the default" and decide what
/// the producing application would have assumed.
///
/// XLSX: `calcPr` in `xl/workbook.xml`. ODS: `table:calculation-settings` in
/// `content.xml` and the default cell style in `styles.xml`.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct DocumentSettings {
    /// Iterative calculation (XLSX `iterate`/`iterateCount`/`iterateDelta`,
    /// ODS `table:iteration`).
    pub iteration: Option<IterationSettings>,
    /// Text comparison is case sensitive (ODS `table:case-sensitive`).
    pub case_sensitive: Option<bool>,
    /// Search criteria must match the whole cell (ODS
    /// `table:search-criteria-must-apply-to-whole-cell`).
    pub search_whole_cell: Option<bool>,
    /// Search criteria are regular expressions (ODS `table:use-regular-expressions`).
    pub use_regular_expressions: Option<bool>,
    /// Search criteria may use `*`/`?`/`~` wildcards (ODS `table:use-wildcards`).
    pub use_wildcards: Option<bool>,
    /// The date of serial 0 as `(year, month, day)` (ODS `table:null-date`; LibreOffice
    /// stores any date). An XLSX states only `date1904`.
    pub null_date: Option<(i64, u32, u32)>,
    /// First year of the 100-year window for two-digit years (ODS `table:null-year`).
    pub null_year: Option<u16>,
    /// Values are rounded to their displayed precision (XLSX
    /// `fullPrecision="0"`, ODS `table:precision-as-shown`).
    pub precision_as_shown: Option<bool>,
    /// How `INDIRECT` reads the reference text (ODS `settings.xml`, `SyntaxStringRef`:
    /// 0 Calc A1 `Sheet.A1`, 2 Excel A1 `Sheet!A1`; an XLSX import stores 2).
    pub string_ref_syntax: Option<u8>,
    /// Calculation on every change (XLSX `calcMode` other than `manual`, ODS `AutoCalculate` of
    /// `settings.xml`); `Some(false)` is the manual mode.
    pub auto_calculate: Option<bool>,
    /// Document language and country as a BCP 47 tag, e.g. `en-US` (ODS: the
    /// default cell style in `styles.xml`).
    pub locale: Option<String>,
}

impl Workbook {
    /// Get a sheet by name
    pub fn get_sheet(&self, name: &str) -> Option<&Sheet> {
        self.sheets.iter().find(|s| s.name == name)
    }

    /// Get sheet name by index
    ///
    /// Returns the sheet name for a given 0-based index.
    /// This is useful for resolving ViolationScope which stores sheet_index.
    pub fn sheet_name_by_index(&self, sheet_index: u16) -> Option<&str> {
        self.sheets
            .iter()
            .find(|s| s.sheet_index == sheet_index)
            .map(|s| s.name.as_str())
    }

    /// Get all sheet names
    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name.as_str()).collect()
    }
}

/// Represents a single worksheet within a workbook
#[derive(Debug, Clone, Default)]
pub struct Sheet {
    /// Name of the sheet
    pub name: String,
    /// Zero-based index representing the order of appearance in the workbook.
    /// This is assigned by the parser based on the sheet's position.
    pub sheet_index: u16,
    /// Map of cell coordinates (row, col) to Cell data. 0-based indexing.
    pub cells: Grid<Cell>,
    /// The used range of the sheet: (max_row, max_col) inclusive. 0-based.
    pub used_range: Option<(u32, u32)>, // (rows, cols)
    /// Hidden column ranges: (start_col, end_col), 0-based, inclusive.
    /// Ranges are in source order and may be adjacent or overlap.
    pub hidden_columns: Vec<(u32, u32)>,
    /// Hidden row ranges: (start_row, end_row), 0-based, inclusive.
    /// Ranges are in source order and may be adjacent or overlap.
    pub hidden_rows: Vec<(u32, u32)>,
    /// Merged cell ranges: (start_row, start_col, end_row, end_col), 0-based, inclusive
    pub merged_cells: Vec<(u32, u32, u32, u32)>,
    /// Error message if there was an error parsing formulas for this sheet
    pub formula_parsing_error: Option<String>,
    /// Internal path to the sheet XML file in the ZIP archive
    pub sheet_path: Option<String>,
    /// Uncompressed size in bytes of the XML that describes this sheet.
    ///
    /// XLSX: size of the sheet's worksheet part (e.g. `xl/worksheets/sheet1.xml`).
    /// ODS: byte span of the sheet's `<table:table>` element inside `content.xml`.
    pub xml_size_bytes: u64,
    /// Number of conditional formatting rules in this sheet
    pub conditional_formatting_count: usize,
    /// Ranges where conditional formatting rules are applied
    pub conditional_formatting_ranges: Vec<String>,
    /// Whether the sheet is visible (not hidden)
    pub visible: bool,
    /// Notes (comments) of the sheet's cells.
    pub comments: usize,
    /// Charts the sheet shows.
    pub charts: usize,
    /// Tables of the sheet: the table objects of an XLSX, the named database ranges of an ODS.
    pub tables: usize,
}

impl Sheet {
    /// Create a new empty sheet with the given name and index
    pub fn new(name: String, sheet_index: u16) -> Self {
        Self {
            name,
            sheet_index,
            cells: Grid::new(),
            used_range: None,
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            formula_parsing_error: None,
            sheet_path: None,
            xml_size_bytes: 0,
            conditional_formatting_count: 0,
            comments: 0,
            charts: 0,
            tables: 0,
            conditional_formatting_ranges: Vec::new(),
            visible: true,
        }
    }

    /// Get a cell at the given position
    pub fn get_cell(&self, row: u32, col: u32) -> Option<&Cell> {
        self.cells.get(&(row, col))
    }

    /// Get all cells with values
    pub fn all_cells(&self) -> impl Iterator<Item = &Cell> {
        self.cells.values()
    }
}

/// Represents a single cell
#[derive(Debug, Clone, Default)]
pub struct Cell {
    /// The typed value of the cell (number, text, boolean, error, or empty).
    pub value: CellValue,
    /// The formula expression (without leading `=`), if any.
    ///
    /// A formula is metadata orthogonal to the value: a cell with `formula: Some("SUM(A1:A10)")`
    /// and `value: Number(42.0)` is a formula cell whose cached result is 42.
    ///
    /// Boxed to minimize `Cell` struct size: `Option<Box<str>>` is 16 bytes vs
    /// `Option<String>`'s 24 bytes, and formula strings are write-once/read-many.
    pub formula: Option<Box<str>>,
    /// Whether this cell is a legacy CSE (Ctrl+Shift+Enter) array formula.
    ///
    /// XLSX: set when `<f t="array">` is present.
    /// ODS: set when `table:number-matrix-columns-spanned` is present.
    pub is_array: bool,
    /// The number format string applied to the cell (e.g., "0.00", "mm/dd/yyyy").
    ///
    /// Interned: cells sharing the same format share one allocation, and the handle is one
    /// pointer.
    pub num_fmt: Option<NumFmt>,
}

/// A number format string, shared by the cells that have it. One pointer wide (an `Arc<str>` is
/// two), which makes the `Cell`, kept by the million, 16 bytes smaller together with the position
/// it no longer repeats: a cell's row and column are those of its place in the sheet's
/// [`Grid`].
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NumFmt(Arc<Box<str>>);

impl std::ops::Deref for NumFmt {
    type Target = str;

    fn deref(&self) -> &str {
        &self.0
    }
}

impl From<&str> for NumFmt {
    fn from(s: &str) -> Self {
        Self(Arc::new(Box::from(s)))
    }
}

impl From<String> for NumFmt {
    fn from(s: String) -> Self {
        Self(Arc::new(s.into_boxed_str()))
    }
}

impl std::fmt::Display for NumFmt {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(self)
    }
}

impl Cell {
    /// Check if the cell contains a formula.
    pub fn is_formula(&self) -> bool {
        self.formula.is_some()
    }

    /// Get the formula expression if this is a formula cell.
    pub fn as_formula(&self) -> Option<&str> {
        self.formula.as_deref()
    }
}

/// Cell value types
#[derive(Debug, Clone, PartialEq, Default)]
pub enum CellValue {
    /// Empty cell
    #[default]
    Empty,
    /// Numeric value (float)
    Number(f64),
    /// Text/String value (interned via `Arc<str>` for memory-efficient sharing)
    Text(Arc<str>),
    /// Boolean value
    Boolean(bool),
    /// Error value (e.g., `#REF!`, `#DIV/0!`, `#N/A`)
    Error(ExcelError),
}

impl CellValue {
    /// Check if the cell contains an error.
    pub fn is_error(&self) -> bool {
        matches!(self, CellValue::Error(_))
    }

    /// Check if the cell is empty.
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    /// Get the error type if this is an error cell.
    pub fn as_error(&self) -> Option<ExcelError> {
        match self {
            CellValue::Error(e) => Some(*e),
            _ => None,
        }
    }
}

#[cfg(test)]
mod size_tests {
    use super::*;

    /// A workbook has millions of cells: what a cell weighs is part of the design (see
    /// `ARCHITECTURE.md`), so growing it is a decision, not a side effect.
    #[test]
    fn a_cell_weighs_56_bytes() {
        assert_eq!(std::mem::size_of::<Cell>(), 56);
        assert_eq!(std::mem::size_of::<Option<NumFmt>>(), 8);
    }
}

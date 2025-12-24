//! Workbook data structures

use std::collections::HashMap;
use std::path::PathBuf;

/// Represents an external workbook reference
#[derive(Debug, Clone)]
pub struct ExternalWorkbook {
    /// 0-based index in the collection (maps to XLSX 1-based index - 1)
    /// For XLSX: index N corresponds to [N+1] in formulas
    /// For ODS: index N corresponds to order of appearance in metadata
    pub index: usize,
    /// Path or URL to the external workbook
    pub path: String,
}

/// Represents a complete workbook
#[derive(Debug, Clone, Default)]
pub struct Workbook {
    pub path: PathBuf,
    pub sheets: Vec<Sheet>,
    pub defined_names: HashMap<String, String>,
    /// List of hidden sheet names
    pub hidden_sheets: Vec<String>,
    /// Whether the workbook contains macros or VBA code
    pub has_macros: bool,
    /// List of external links (URLs or workbook paths) found in metadata
    /// Deprecated: Use external_workbooks for structured access with indices
    pub external_links: Vec<String>,
    /// External workbooks referenced by this workbook (0-indexed collection)
    /// For XLSX: index N corresponds to [N+1] in formulas
    /// For ODS: index N corresponds to order of appearance in metadata
    pub external_workbooks: Vec<ExternalWorkbook>,
}

impl Workbook {
    /// Get a sheet by name
    pub fn get_sheet(&self, name: &str) -> Option<&Sheet> {
        self.sheets.iter().find(|s| s.name == name)
    }

    /// Get all sheet names
    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name.as_str()).collect()
    }
}

/// Represents a worksheet
#[derive(Debug, Clone, Default)]
pub struct Sheet {
    pub name: String,
    pub cells: HashMap<(u32, u32), Cell>,
    pub used_range: Option<(u32, u32)>, // (rows, cols)
    /// List of hidden column indices (0-based)
    pub hidden_columns: Vec<u32>,
    /// List of hidden row indices (0-based)
    pub hidden_rows: Vec<u32>,
    /// Merged cell ranges: (start_row, start_col, end_row, end_col)
    pub merged_cells: Vec<(u32, u32, u32, u32)>,
    /// Error message if there was an error parsing formulas for this sheet
    pub formula_parsing_error: Option<String>,
    /// Internal path to the sheet XML file in the ZIP archive
    pub sheet_path: Option<String>,
    /// Number of conditional formatting rules in this sheet
    pub conditional_formatting_count: usize,
    /// Ranges where conditional formatting rules are applied
    pub conditional_formatting_ranges: Vec<String>,
    pub visible: bool,
}

impl Sheet {
    pub fn new(name: String) -> Self {
        Self {
            name,
            cells: HashMap::new(),
            used_range: None,
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            formula_parsing_error: None,
            sheet_path: None,
            conditional_formatting_count: 0,
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

    /// Get cells in a specific column
    pub fn cells_in_column(&self, col: u32) -> impl Iterator<Item = &Cell> {
        self.cells.values().filter(move |c| c.col == col)
    }

    /// Get cells in a specific row
    pub fn cells_in_row(&self, row: u32) -> impl Iterator<Item = &Cell> {
        self.cells.values().filter(move |c| c.row == row)
    }

    /// Get the last cell with actual data (bottom-right corner of data range)
    pub fn last_data_cell(&self) -> Option<(u32, u32)> {
        let non_empty_cells: Vec<_> = self
            .cells
            .values()
            .filter(|c| !matches!(c.value, CellValue::Empty))
            .collect();

        if non_empty_cells.is_empty() {
            return None;
        }

        // Find the maximum row and maximum column independently
        let max_row = non_empty_cells.iter().map(|c| c.row).max()?;
        let max_col = non_empty_cells.iter().map(|c| c.col).max()?;

        Some((max_row, max_col))
    }
}

/// Represents a single cell
#[derive(Debug, Clone, Default)]
pub struct Cell {
    pub row: u32,
    pub col: u32,
    pub value: CellValue,
    pub num_fmt: Option<String>,
}

/// Cell value types
#[derive(Debug, Clone, PartialEq, Default)]
pub enum CellValue {
    #[default]
    Empty,
    Number(f64),
    Text(String),
    Boolean(bool),
    Formula {
        formula: String,
        cached_error: Option<String>,
    },
}

impl CellValue {
    /// Check if the cell contains an error
    pub fn is_error(&self) -> bool {
        matches!(
            self,
            CellValue::Formula {
                cached_error: Some(_),
                ..
            }
        )
    }

    /// Check if the cell is empty
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    /// Check if the cell contains a formula
    pub fn is_formula(&self) -> bool {
        matches!(self, CellValue::Formula { .. })
    }

    /// Get the error value if this is an error cell
    pub fn as_error(&self) -> Option<&str> {
        match self {
            CellValue::Formula {
                cached_error: Some(e),
                ..
            } => Some(e),
            _ => None,
        }
    }

    /// Get the formula if this is a formula cell
    pub fn as_formula(&self) -> Option<&str> {
        match self {
            CellValue::Formula { formula, .. } => Some(formula),
            _ => None,
        }
    }

    /// Create a formula cell without error
    pub fn formula(f: impl Into<String>) -> Self {
        CellValue::Formula {
            formula: f.into(),
            cached_error: None,
        }
    }

    /// Create a formula cell with cached error
    pub fn formula_with_error(f: impl Into<String>, error: impl Into<String>) -> Self {
        CellValue::Formula {
            formula: f.into(),
            cached_error: Some(error.into()),
        }
    }
}

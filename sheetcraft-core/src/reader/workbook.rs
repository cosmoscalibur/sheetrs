//! Workbook data structures

use std::collections::HashMap;
use std::path::PathBuf;

/// Represents a complete workbook
#[derive(Debug, Clone)]
pub struct Workbook {
    pub path: PathBuf,
    pub sheets: Vec<Sheet>,
    pub defined_names: HashMap<String, String>,
    /// List of hidden sheet names
    pub hidden_sheets: Vec<String>,
    /// Whether the workbook contains macros or VBA code
    pub has_macros: bool,
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
#[derive(Debug, Clone)]
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
}

impl Sheet {
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
#[derive(Debug, Clone)]
pub struct Cell {
    pub row: u32,
    pub col: u32,
    pub value: CellValue,
    pub num_fmt: Option<String>,
}

/// Cell value types
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Empty,
    Number(f64),
    Text(String),
    Boolean(bool),
    Error(String),
    Formula(String),
}

impl CellValue {
    /// Check if the cell contains an error
    pub fn is_error(&self) -> bool {
        matches!(self, CellValue::Error(_))
    }

    /// Check if the cell is empty
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    /// Check if the cell contains a formula
    pub fn is_formula(&self) -> bool {
        matches!(self, CellValue::Formula(_))
    }

    /// Get the error value if this is an error cell
    pub fn as_error(&self) -> Option<&str> {
        match self {
            CellValue::Error(e) => Some(e),
            _ => None,
        }
    }

    /// Get the formula if this is a formula cell
    pub fn as_formula(&self) -> Option<&str> {
        match self {
            CellValue::Formula(f) => Some(f),
            _ => None,
        }
    }
}

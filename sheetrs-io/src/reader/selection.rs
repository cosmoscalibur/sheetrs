//! What to read of a workbook: which sheets, which cells of them and which properties of each
//! cell.
//!
//! A [`Selection`] is handed to the reader, which does less work for it instead of giving
//! everything to the receiver: a sheet nobody asked for is not parsed, the reading of a sheet
//! stops after the last row asked for, the cells outside the windows are skipped, and the
//! properties that were not asked for are not built (the text of a formula, the format of a
//! number). A selective read does not report whether a formula is an array formula: `is_array` is
//! left `false`. The memory of a read is that of what was asked.
//!
//! Rows are read in order: the reading of a sheet ends at the first row after the last one a
//! window reaches. An ODS has them in order by its structure. An XLSX is expected to have them
//! in ascending order, as Excel requires; in a file that does not, the cells that follow the stop
//! are not read.

use sheetrs_formula::addr::same_name;

/// A rectangle of cells of a sheet: rows and columns, 0-based, both ends included.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Window {
    pub first_row: u32,
    pub last_row: u32,
    pub first_col: u32,
    pub last_col: u32,
}

impl Window {
    /// Every cell of a sheet.
    pub const ALL: Self = Self {
        first_row: 0,
        last_row: u32::MAX,
        first_col: 0,
        last_col: u32::MAX,
    };

    /// The cells between two corners, in any order.
    #[must_use]
    pub fn new(row_a: u32, col_a: u32, row_b: u32, col_b: u32) -> Self {
        Self {
            first_row: row_a.min(row_b),
            last_row: row_a.max(row_b),
            first_col: col_a.min(col_b),
            last_col: col_a.max(col_b),
        }
    }

    /// Whether the cell at `(row, col)` is inside.
    #[must_use]
    pub fn contains(&self, row: u32, col: u32) -> bool {
        (self.first_row..=self.last_row).contains(&row)
            && (self.first_col..=self.last_col).contains(&col)
    }
}

/// The properties of a cell that the reader builds. What is not asked for is left empty in the
/// [`Cell`](super::Cell) it hands over.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Properties {
    /// The cached value (`Cell::value`).
    pub value: bool,
    /// The text of the formula (`Cell::formula`).
    pub formula: bool,
    /// The number format (`Cell::num_fmt`).
    pub number_format: bool,
}

impl Properties {
    /// All of them.
    pub const ALL: Self = Self {
        value: true,
        formula: true,
        number_format: true,
    };

    /// Only the value.
    pub const VALUE: Self = Self {
        value: true,
        formula: false,
        number_format: false,
    };
}

/// The sheets to read, the windows of each, and the properties wanted of their cells.
#[derive(Debug, Clone)]
pub struct Selection {
    sheets: Vec<(String, Vec<Window>)>,
    properties: Properties,
}

impl Selection {
    /// A selection of nothing yet, for the given properties.
    #[must_use]
    pub fn new(properties: Properties) -> Self {
        Self {
            sheets: Vec::new(),
            properties,
        }
    }

    /// Adds a window of the sheet called `sheet` (the name is matched without regard to case,
    /// as references are).
    #[must_use]
    pub fn window(mut self, sheet: impl Into<String>, window: Window) -> Self {
        let sheet = sheet.into();
        match self.sheets.iter_mut().find(|(n, _)| same_name(n, &sheet)) {
            Some((_, windows)) => windows.push(window),
            None => self.sheets.push((sheet, vec![window])),
        }
        self
    }

    /// The properties wanted.
    #[must_use]
    pub fn properties(&self) -> Properties {
        self.properties
    }

    /// How many sheets are selected.
    #[must_use]
    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// The part of the selection that is of `sheet`, if the sheet is selected.
    pub(crate) fn of(&self, sheet: &str) -> Option<SheetSelection<'_>> {
        self.sheets
            .iter()
            .find(|(n, _)| same_name(n, sheet))
            .map(|(_, windows)| SheetSelection::new(windows))
    }
}

/// The windows of one sheet, with the bounds that the readers use to skip and to stop.
#[derive(Debug, Clone, Copy)]
pub(crate) struct SheetSelection<'a> {
    windows: &'a [Window],
    /// The first row any window reaches: the rows before it hold nothing wanted.
    pub first_row: u32,
    /// The last row any window reaches: the reading of the sheet ends after it.
    pub last_row: u32,
    first_col: u32,
    last_col: u32,
}

impl<'a> SheetSelection<'a> {
    fn new(windows: &'a [Window]) -> Self {
        Self {
            windows,
            first_row: windows.iter().map(|w| w.first_row).min().unwrap_or(0),
            last_row: windows.iter().map(|w| w.last_row).max().unwrap_or(0),
            first_col: windows.iter().map(|w| w.first_col).min().unwrap_or(0),
            last_col: windows.iter().map(|w| w.last_col).max().unwrap_or(0),
        }
    }

    /// Whether the cell at `(row, col)` is in a window.
    pub fn contains(&self, row: u32, col: u32) -> bool {
        self.windows.iter().any(|w| w.contains(row, col))
    }

    /// Whether a block of `rows` rows from `row` and `cols` columns from `col` (a cell, or the
    /// run of cells that a repeat attribute stands for) may hold anything wanted: a test by the
    /// bounds of the windows, no stricter than [`contains`](Self::contains).
    pub fn may_contain_block(&self, row: u32, rows: u32, col: u32, cols: u32) -> bool {
        row <= self.last_row
            && row.saturating_add(rows) > self.first_row
            && col <= self.last_col
            && col.saturating_add(cols) > self.first_col
    }
}

/// A sheet of the workbook, as the selective readers list them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetInfo {
    pub name: String,
    /// The index the sheet's cells carry when they are handed to the receiver.
    pub index: u16,
    pub state: super::SheetState,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn windows_of_the_same_sheet_join_and_names_match_without_case() {
        let selection = Selection::new(Properties::VALUE)
            .window("Data", Window::new(2, 1, 5, 3))
            .window("DATA", Window::new(10, 0, 12, 0))
            .window("Other", Window::ALL);
        assert_eq!(selection.sheet_count(), 2);
        let data = selection.of("data").unwrap();
        assert_eq!((data.first_row, data.last_row), (2, 12));
        assert!(data.contains(3, 2) && data.contains(11, 0));
        assert!(!data.contains(8, 2), "between the two windows");
        assert!(data.may_contain_block(8, 1, 2, 1), "inside their bounds");
        assert!(selection.of("Missing").is_none());
    }

    #[test]
    fn a_window_takes_its_corners_in_any_order() {
        let w = Window::new(5, 4, 1, 2);
        assert_eq!(
            (w.first_row, w.last_row, w.first_col, w.last_col),
            (1, 5, 2, 4)
        );
    }

    #[test]
    fn blocks_are_tested_against_the_bounds() {
        let selection = Selection::new(Properties::ALL).window("S", Window::new(10, 5, 20, 8));
        let s = selection.of("S").unwrap();
        assert!(
            s.may_contain_block(10, 1, 0, 6),
            "a run that reaches column 5"
        );
        assert!(
            !s.may_contain_block(10, 1, 0, 5),
            "a run that stops at column 4"
        );
        assert!(!s.may_contain_block(10, 1, 9, 3), "starts after column 8");
        assert!(!s.may_contain_block(9, 1, 5, 1), "a row before");
        assert!(
            s.may_contain_block(8, 3, 0, 6),
            "rows 8 to 10 and columns 0 to 5 reach it"
        );
        assert!(!s.may_contain_block(21, 4, 5, 1), "below the last row");
    }
}

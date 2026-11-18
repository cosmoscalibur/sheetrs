//! Workbook model: sheets of cells, as read from and written to files.

use std::collections::BTreeMap;
use std::sync::Arc;

use crate::coord::CellPos;

/// A spreadsheet error value, identified by the text a file stores for it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ErrorKind {
    /// `#NULL!`
    Null,
    /// `#DIV/0!`
    DivZero,
    /// `#VALUE!`
    Value,
    /// `#REF!`
    Ref,
    /// `#NAME?`
    Name,
    /// `#NUM!`
    Num,
    /// `#N/A`
    Na,
    /// Any other error text, kept as stored.
    Other(Arc<str>),
}

impl ErrorKind {
    /// Reads the error stored as `text`; text that is not a known error is kept in [`ErrorKind::Other`].
    #[must_use]
    pub fn from_text(text: &str) -> Self {
        match text {
            "#NULL!" => Self::Null,
            "#DIV/0!" => Self::DivZero,
            "#VALUE!" => Self::Value,
            "#REF!" => Self::Ref,
            "#NAME?" => Self::Name,
            "#NUM!" => Self::Num,
            "#N/A" => Self::Na,
            other => Self::Other(other.into()),
        }
    }

    /// The text files store for this error.
    #[must_use]
    pub fn as_text(&self) -> &str {
        match self {
            Self::Null => "#NULL!",
            Self::DivZero => "#DIV/0!",
            Self::Value => "#VALUE!",
            Self::Ref => "#REF!",
            Self::Name => "#NAME?",
            Self::Num => "#NUM!",
            Self::Na => "#N/A",
            Self::Other(text) => text,
        }
    }
}

/// The value of a cell: the constant it holds, or the cached result of its formula.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    /// No value.
    Empty,
    /// A number; dates are numbers too.
    Number(f64),
    /// A boolean.
    Bool(bool),
    /// Text.
    Text(Arc<str>),
    /// An error value.
    Error(ErrorKind),
}

/// The formula of a cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Formula {
    /// Formula text as stored in the file, without a leading `=`.
    Text(String),
    /// A cell that reuses the formula of a shared group; the anchor cell is
    /// given by [`Sheet::shared_anchor`], which is `None` when the file does
    /// not contain the group's anchor.
    Shared {
        /// The group the cell belongs to.
        index: u32,
    },
}

/// One cell: its value and, when it holds a formula, the formula.
#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    /// The constant or cached result.
    pub value: Value,
    /// The formula, when the cell has one.
    pub formula: Option<Formula>,
    /// Whether `value` is a cache that `LibreOffice` does not trust on load and recalculates.
    pub stale: bool,
}

/// Whether a sheet is shown to the user.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Visibility {
    /// Shown.
    Visible,
    /// Hidden; the user can unhide it.
    Hidden,
    /// Hidden; the user cannot unhide it from the UI.
    VeryHidden,
}

/// A sheet: a name and its non-empty cells.
#[derive(Debug, Clone)]
pub struct Sheet {
    name: String,
    visibility: Visibility,
    cells: BTreeMap<CellPos, Cell>,
    shared_anchors: BTreeMap<u32, CellPos>,
}

impl Sheet {
    /// Creates an empty, visible sheet.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            visibility: Visibility::Visible,
            cells: BTreeMap::new(),
            shared_anchors: BTreeMap::new(),
        }
    }

    /// The sheet name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Whether the sheet is shown.
    #[must_use]
    pub fn visibility(&self) -> Visibility {
        self.visibility
    }

    /// Sets whether the sheet is shown.
    pub fn set_visibility(&mut self, visibility: Visibility) {
        self.visibility = visibility;
    }

    /// Stores `cell` at `pos`, replacing any previous cell.
    pub fn set(&mut self, pos: CellPos, cell: Cell) {
        self.cells.insert(pos, cell);
    }

    /// The cell at `pos`, if it is not empty.
    #[must_use]
    pub fn get(&self, pos: CellPos) -> Option<&Cell> {
        self.cells.get(&pos)
    }

    /// The non-empty cells in row-major order.
    pub fn cells(&self) -> impl Iterator<Item = (CellPos, &Cell)> {
        self.cells.iter().map(|(pos, cell)| (*pos, cell))
    }

    /// The number of non-empty cells.
    #[must_use]
    pub fn len(&self) -> usize {
        self.cells.len()
    }

    /// Whether the sheet has no cells.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty()
    }

    /// Records the cell that carries the formula text of shared group `index`.
    pub fn set_shared_anchor(&mut self, index: u32, anchor: CellPos) {
        self.shared_anchors.insert(index, anchor);
    }

    /// The cell that carries the formula text of shared group `index`.
    #[must_use]
    pub fn shared_anchor(&self, index: u32) -> Option<CellPos> {
        self.shared_anchors.get(&index).copied()
    }
}

/// How serial dates count: from 1899-12-30 or from 1904-01-01.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DateSystem {
    /// Serial 0 is 1899-12-30.
    #[default]
    Date1900,
    /// Serial 0 is 1904-01-01.
    Date1904,
}

/// An ordered list of sheets.
#[derive(Debug, Clone, Default)]
pub struct Workbook {
    sheets: Vec<Sheet>,
    date_system: DateSystem,
}

impl Workbook {
    /// Creates a workbook without sheets.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// The sheets, in workbook order.
    #[must_use]
    pub fn sheets(&self) -> &[Sheet] {
        &self.sheets
    }

    /// The first sheet called `name`, compared exactly (case-sensitive).
    #[must_use]
    pub fn sheet(&self, name: &str) -> Option<&Sheet> {
        self.sheets.iter().find(|sheet| sheet.name == name)
    }

    /// Appends a sheet.
    pub fn push_sheet(&mut self, sheet: Sheet) {
        self.sheets.push(sheet);
    }

    /// How the workbook counts serial dates.
    #[must_use]
    pub fn date_system(&self) -> DateSystem {
        self.date_system
    }

    /// Sets how the workbook counts serial dates.
    pub fn set_date_system(&mut self, date_system: DateSystem) {
        self.date_system = date_system;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn number(v: f64) -> Cell {
        Cell {
            value: Value::Number(v),
            formula: None,
            stale: false,
        }
    }

    #[test]
    fn cells_iterate_in_row_major_order() {
        let mut sheet = Sheet::new("S1");
        sheet.set(CellPos { row: 1, col: 0 }, number(3.0));
        sheet.set(CellPos { row: 0, col: 2 }, number(2.0));
        sheet.set(CellPos { row: 0, col: 0 }, number(1.0));
        let order: Vec<(u32, u32)> = sheet.cells().map(|(pos, _)| (pos.row, pos.col)).collect();
        assert_eq!(order, vec![(0, 0), (0, 2), (1, 0)]);
        assert_eq!(sheet.len(), 3);
        assert!(!sheet.is_empty());
    }

    #[test]
    fn setting_a_cell_replaces_the_previous_one() {
        let mut sheet = Sheet::new("S1");
        let pos = CellPos { row: 0, col: 0 };
        sheet.set(pos, number(1.0));
        sheet.set(pos, number(2.0));
        assert_eq!(sheet.get(pos), Some(&number(2.0)));
        assert_eq!(sheet.len(), 1);
    }

    #[test]
    fn error_text_round_trips() {
        for text in [
            "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!",
        ] {
            assert_eq!(ErrorKind::from_text(text).as_text(), text);
        }
        assert_eq!(
            ErrorKind::from_text("#SPILL!"),
            ErrorKind::Other("#SPILL!".into())
        );
    }

    #[test]
    fn workbook_finds_sheets_by_name() {
        let mut workbook = Workbook::new();
        workbook.push_sheet(Sheet::new("S1"));
        workbook.push_sheet(Sheet::new("S2"));
        assert_eq!(workbook.sheet("S2").map(Sheet::name), Some("S2"));
        assert!(workbook.sheet("S3").is_none());
        assert_eq!(workbook.sheets().len(), 2);
        assert_eq!(workbook.date_system(), DateSystem::Date1900);
    }
}

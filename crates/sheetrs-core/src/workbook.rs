//! Workbook model: sheets of cells, as read from and written to files.

use std::sync::Arc;

use crate::store::{Sheet, SheetBuilder, SheetView, Tables};

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
    /// given by [`SheetView::shared_anchor`], which is `None` when the file does
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

/// How serial dates count: from 1899-12-30 or from 1904-01-01.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DateSystem {
    /// Serial 0 is 1899-12-30.
    #[default]
    Date1900,
    /// Serial 0 is 1904-01-01.
    Date1904,
}

/// An ordered list of sheets, with the tables of strings and formulas they share.
#[derive(Debug, Clone, Default)]
pub struct Workbook {
    pub(crate) sheets: Vec<Sheet>,
    pub(crate) tables: Tables,
    date_system: DateSystem,
}

impl Workbook {
    /// Creates a workbook without sheets.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// The sheets, in workbook order.
    pub fn sheets(&self) -> impl ExactSizeIterator<Item = SheetView<'_>> {
        self.sheets
            .iter()
            .map(|sheet| SheetView::new(&self.tables, sheet))
    }

    /// The first sheet called `name`, compared exactly (case-sensitive).
    #[must_use]
    pub fn sheet(&self, name: &str) -> Option<SheetView<'_>> {
        self.sheets
            .iter()
            .find(|sheet| sheet.name() == name)
            .map(|sheet| SheetView::new(&self.tables, sheet))
    }

    /// Starts a sheet called `name`; it joins the workbook when the builder is finished.
    pub fn sheet_builder(&mut self, name: impl Into<String>) -> SheetBuilder<'_> {
        SheetBuilder::new(self, name.into())
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
    fn workbooks_use_the_1900_date_system_by_default() {
        assert_eq!(Workbook::new().date_system(), DateSystem::Date1900);
    }
}

//! Workbook model: sheets of cells, as read from and written to files.

use std::sync::Arc;

use crate::opaque::OpaquePackage;
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
    /// Formula text as the file stores it, in the grammar of its format.
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

/// How serial dates count: from 1899-12-30, from 1904-01-01 or from a date of the file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DateSystem {
    /// Serial 0 is 1899-12-30.
    #[default]
    Date1900,
    /// Serial 0 is 1904-01-01.
    Date1904,
    /// Serial 0 is another date, which some formats let the file choose.
    Custom {
        /// Days from 1970-01-01 to the date of serial 0, negative before it.
        zero: i64,
    },
}

/// The grammar in which the formulas of a workbook are written, which is the one of the format it
/// was read from.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FormulaGrammar {
    /// Office Open XML: the text of an `<f>` element, without `=`.
    #[default]
    Ooxml,
    /// `OpenFormula` as ODS writes it, with its namespace prefix and `=`.
    OpenFormula,
    /// The syntax `LibreOffice` itself uses, with `=`, as a text file holds it.
    Native,
}

/// A document that the formulas of a workbook refer to.
///
/// The value of such a reference is the one the file caches; it is never recalculated, because
/// the document is not opened.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalLink {
    /// Where the document is, as the file says: an address, or a path relative to the workbook.
    pub target: String,
}

/// What kind of code a [`MacroModule`] holds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MacroKind {
    /// A module of a Basic library, in the form that `LibreOffice` keeps it.
    Basic,
    /// A module of a VBA project.
    Vba,
}

/// A module of macro code that a workbook holds. The code is listed, never run.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MacroModule {
    pub kind: MacroKind,
    /// The library or project the module is in.
    pub library: String,
    pub name: String,
}

/// An ordered list of sheets, with the tables of strings and formulas they share.
#[derive(Debug, Clone, Default)]
pub struct Workbook {
    pub(crate) sheets: Vec<Sheet>,
    pub(crate) tables: Tables,
    date_system: DateSystem,
    grammar: FormulaGrammar,
    external_links: Vec<ExternalLink>,
    /// The targets of `external_links`, to tell a repeated one at once.
    external_targets: std::collections::HashSet<String>,
    macro_modules: Vec<MacroModule>,
    opaque: Option<OpaquePackage>,
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

    /// What the package the workbook was read from holds besides the model, if the workbook was
    /// read from a package that carries it.
    #[must_use]
    pub fn opaque(&self) -> Option<&OpaquePackage> {
        self.opaque.as_ref()
    }

    /// Sets what the workbook carries from its package.
    pub fn set_opaque(&mut self, package: OpaquePackage) {
        self.opaque = Some(package);
    }

    /// The grammar of the formulas of the workbook.
    #[must_use]
    pub fn formula_grammar(&self) -> FormulaGrammar {
        self.grammar
    }

    /// Sets the grammar of the formulas of the workbook.
    pub fn set_formula_grammar(&mut self, grammar: FormulaGrammar) {
        self.grammar = grammar;
    }

    /// Sets how the workbook counts serial dates.
    pub fn set_date_system(&mut self, date_system: DateSystem) {
        self.date_system = date_system;
    }

    /// The documents that formulas of the workbook refer to, each once, in the order they were
    /// found.
    #[must_use]
    pub fn external_links(&self) -> &[ExternalLink] {
        &self.external_links
    }

    /// The modules of macro code that the workbook holds, in the order they were found.
    #[must_use]
    pub fn macro_modules(&self) -> &[MacroModule] {
        &self.macro_modules
    }

    /// Records a module of macro code.
    pub fn add_macro_module(&mut self, module: MacroModule) {
        self.macro_modules.push(module);
    }

    /// Records that a formula refers to the document at `target`, unless it is recorded.
    pub fn add_external_link(&mut self, target: &str) {
        if self.external_targets.insert(target.to_owned()) {
            self.external_links.push(ExternalLink {
                target: target.to_owned(),
            });
        }
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

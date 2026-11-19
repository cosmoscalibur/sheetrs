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
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
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

/// A name that stands for an expression, such as a range, that formulas can use.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedExpression {
    /// The name, which is unique within its scope without regard to case.
    pub name: String,
    /// The index of the sheet the name belongs to; `None` for a name of the workbook.
    pub scope: Option<usize>,
    /// The expression in the grammar of the workbook (see [`FormulaGrammar`]), without a
    /// leading `=` and without surrounding spaces.
    pub expression: String,
    /// Whether the name is kept out of the lists offered to the user.
    pub hidden: bool,
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
    names: Vec<NamedExpression>,
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

    /// The names of the workbook and of its sheets, in the order they were defined.
    #[must_use]
    pub fn names(&self) -> &[NamedExpression] {
        &self.names
    }

    /// Defines a name as `LibreOffice` loads it: when its scope already has a name that is the same
    /// without regard to case, the new one is stored with `_0` after it, or `_1` and so on until
    /// it is unique. Returns the name as stored.
    pub fn define_name(&mut self, mut name: NamedExpression) -> &str {
        let taken = |names: &[NamedExpression], candidate: &str| {
            names.iter().any(|other| {
                other.scope == name.scope && other.name.to_lowercase() == candidate.to_lowercase()
            })
        };
        if taken(&self.names, &name.name) {
            let base = name.name.clone();
            let mut number = 0_u32;
            loop {
                let candidate = format!("{base}_{number}");
                if !taken(&self.names, &candidate) {
                    name.name = candidate;
                    break;
                }
                number += 1;
            }
        }
        self.names.push(name);
        self.names.last().map_or("", |stored| stored.name.as_str())
    }

    /// Defines a name as `LibreOffice` loads it from an ODS document: when its scope already has a
    /// name that is the same without regard to case, the new definition replaces it.
    pub fn define_name_replacing(&mut self, name: NamedExpression) {
        let lowered = name.name.to_lowercase();
        self.names
            .retain(|other| other.scope != name.scope || other.name.to_lowercase() != lowered);
        self.names.push(name);
    }

    /// A number that is the same for two workbooks that have the same sheets, cells, formulas,
    /// names and date system, and that differs, but for a collision, when they do not. It is
    /// not stable between versions.
    #[must_use]
    pub fn fingerprint(&self) -> u64 {
        use std::hash::{Hash, Hasher};
        let mut hasher = rustc_hash::FxHasher::default();
        self.date_system.hash(&mut hasher);
        (self.grammar as u8).hash(&mut hasher);
        for sheet in self.sheets() {
            sheet.name().hash(&mut hasher);
            (sheet.visibility() as u8).hash(&mut hasher);
            for (pos, cell) in sheet.cells_by_column() {
                (pos.row, pos.col, cell.stale).hash(&mut hasher);
                match cell.value {
                    crate::ValueRef::Empty => 0_u8.hash(&mut hasher),
                    crate::ValueRef::Number(number) => (1_u8, number.to_bits()).hash(&mut hasher),
                    crate::ValueRef::Bool(flag) => (2_u8, flag).hash(&mut hasher),
                    crate::ValueRef::Text(text) => (3_u8, text).hash(&mut hasher),
                    crate::ValueRef::Error(kind) => (4_u8, kind.as_text()).hash(&mut hasher),
                }
                match cell.formula {
                    None => 0_u8.hash(&mut hasher),
                    Some(crate::FormulaRef::Text(text)) => (1_u8, text).hash(&mut hasher),
                    Some(crate::FormulaRef::Shared { index }) => (2_u8, index).hash(&mut hasher),
                }
            }
            for (first, last) in sheet.array_areas() {
                (first, last).hash(&mut hasher);
            }
        }
        for name in &self.names {
            (&name.name, name.scope, &name.expression, name.hidden).hash(&mut hasher);
        }
        hasher.finish()
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
    fn names_that_repeat_without_regard_to_case_are_numbered() {
        let mut workbook = Workbook::new();
        let mut define = |name: &str, scope| {
            workbook
                .define_name(NamedExpression {
                    name: name.to_owned(),
                    scope,
                    expression: "1".to_owned(),
                    hidden: false,
                })
                .to_owned()
        };
        assert_eq!(define("Bb", None), "Bb");
        assert_eq!(define("bB", None), "bB_0");
        assert_eq!(define("BB", None), "BB_1");
        assert_eq!(define("Bb", Some(0)), "Bb");
    }

    #[test]
    fn a_name_that_repeats_replaces_the_earlier_one_when_asked() {
        let mut workbook = Workbook::new();
        for (name, expression) in [("Bb", "1"), ("bB", "2")] {
            workbook.define_name_replacing(NamedExpression {
                name: name.to_owned(),
                scope: None,
                expression: expression.to_owned(),
                hidden: false,
            });
        }
        assert_eq!(workbook.names().len(), 1);
        assert_eq!(workbook.names()[0].name, "bB");
    }

    fn workbook_with(number: f64, name: &str) -> Workbook {
        let mut workbook = Workbook::new();
        let mut sheet = workbook.sheet_builder("S1");
        let cell = Cell {
            value: Value::Number(number),
            formula: None,
            stale: false,
        };
        sheet
            .set(crate::CellPos { row: 0, col: 0 }, &cell)
            .expect("one cell fits");
        sheet.finish();
        workbook.define_name_replacing(NamedExpression {
            name: name.to_owned(),
            scope: None,
            expression: "1".to_owned(),
            hidden: false,
        });
        workbook
    }

    #[test]
    fn the_fingerprint_follows_cells_and_names() {
        let base = workbook_with(1.0, "N").fingerprint();
        assert_eq!(base, workbook_with(1.0, "N").fingerprint());
        assert_ne!(base, workbook_with(2.0, "N").fingerprint());
        assert_ne!(base, workbook_with(1.0, "M").fingerprint());
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
    fn workbooks_use_the_1900_date_system_by_default() {
        assert_eq!(Workbook::new().date_system(), DateSystem::Date1900);
    }
}

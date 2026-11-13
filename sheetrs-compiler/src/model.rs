//! In-memory, format-independent workbook: what every frontend produces and
//! the compiler consumes. Formulas stay as text until compilation, once all
//! sheet and defined names are known.

use crate::cellstore::CellStore;
use sheetrs_engine::addr::{self, CellRef};
use sheetrs_engine::options::{DateSystem, Iteration, Search};
use sheetrs_engine::value::Value;
use std::collections::HashMap;

#[derive(Debug, Clone, PartialEq)]
pub struct ModelCell {
    /// Constant value, or the cached result of the formula.
    pub value: Value,
    /// Formula text without the leading `=`, in Excel syntax. `Some("")` is a formula whose text
    /// was released (see [`ModelCell::release_formula`]): the cell still is a formula.
    pub formula: Option<String>,
    /// Array (Ctrl+Shift+Enter) formula: its operators work element-wise.
    pub array: bool,
}

impl ModelCell {
    /// Frees the text of the formula, as `compile` does of a model it owns once every formula is
    /// lowered: after that only whether the cell is a formula is read (`formula.is_some()`).
    pub fn release_formula(&mut self) {
        if let Some(text) = &mut self.formula {
            *text = String::new();
        }
    }
}

/// Calculation settings stored in the document, resolved into engine terms.
/// `None`: not stated, the engine default applies.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ModelSettings {
    /// Iterative calculation, when the document switches it on.
    pub iteration: Option<Iteration>,
    pub search: Option<Search>,
    pub two_digit_year_start: Option<u16>,
    /// `INDIRECT` reads `Sheet!A1` (Excel A1, what an XLSX import sets) and not `Sheet.A1`
    /// (Calc A1, a document created in `LibreOffice`). `None`: Excel's.
    pub string_ref_excel: Option<bool>,
    /// The document is in manual calculation mode (it calculates only on request).
    pub manual_calc: bool,
    /// Document language tag (`es-CO`): an ODS says it, an XLSX does not.
    pub locale: Option<String>,
    /// The locale given to the workbook (`Book::with_locale`): it is the locale of the run unless the
    /// configuration states another.
    pub book_locale: Option<String>,
    /// The locale of the system the program runs on, which an XLSX (that declares none) takes.
    pub system_locale: Option<String>,
}

#[derive(Debug, Clone)]
pub struct Model {
    pub sheets: Vec<String>,
    pub cells: CellStore,
    /// The cells of an array formula that are not its top-left one: how far below and to the right
    /// of that cell each one sits (`(rows, columns)`), for it to show its own element of the
    /// result. A model that is not read from a file holds none, and every cell of an array formula
    /// then shows the first element.
    pub array_members: HashMap<CellRef, (u32, u32)>,
    /// Defined names: name → formula text (usually a reference).
    pub names: HashMap<String, String>,
    pub date_system: DateSystem,
    pub settings: ModelSettings,
    /// Built by a file loader: the document defines date system, iteration,
    /// search settings and reference syntax, and a `Spec` may not override them.
    pub from_file: bool,
    /// The formulas come from an XLSX (`LibreOffice` reads a bare `#REF!` in one as an
    /// unknown name).
    pub from_xlsx: bool,
    /// The file name of the workbook, with its extension (`CELL("filename")`).
    pub file_name: Option<String>,
    /// The file contains macros (VBA, Basic, scripts). They are never run.
    pub macros: bool,
    /// The document properties the file declares (`title`, `creator`, `created`…), in file order.
    pub properties: Vec<(String, String)>,
    /// The file the model was read from: its format and hash.
    pub source: sheetrs_engine::ir::SourceInfo,
    /// Hidden rows `(sheet, first, last)`, 0-based inclusive.
    pub hidden_rows: Vec<(u16, u32, u32)>,
    pub diagnostics: Vec<String>,
    /// Document options the engine does not support: a build refuses the workbook.
    pub unsupported_options: Vec<String>,
}

impl Default for Model {
    fn default() -> Self {
        Self::new()
    }
}

impl Model {
    #[must_use]
    pub fn new() -> Self {
        Self {
            sheets: Vec::new(),
            cells: CellStore::default(),
            array_members: HashMap::new(),
            names: HashMap::new(),
            // LibreOffice numbers dates from 1899-12-30 on the real calendar.
            date_system: DateSystem::Null1899,
            settings: ModelSettings::default(),
            from_file: false,
            from_xlsx: false,
            file_name: None,
            macros: false,
            properties: Vec::new(),
            source: sheetrs_engine::ir::SourceInfo::default(),
            hidden_rows: Vec::new(),
            diagnostics: Vec::new(),
            unsupported_options: Vec::new(),
        }
    }

    pub fn add_sheet(&mut self, name: &str) -> u16 {
        self.sheets.push(name.to_string());
        (self.sheets.len() - 1) as u16
    }

    fn at(sheet: u16, a1: &str) -> CellRef {
        let area = addr::parse_area(a1).unwrap_or_else(|| panic!("bad address `{a1}`"));
        CellRef::new(sheet, area.r0, area.c0)
    }

    /// Set a constant, addressing the cell in A1 notation.
    pub fn set(&mut self, sheet: u16, a1: &str, value: Value) -> &mut Self {
        let at = Self::at(sheet, a1);
        self.cells.insert(
            at,
            ModelCell {
                value,
                formula: None,
                array: false,
            },
        );
        self
    }

    pub fn set_formula(&mut self, sheet: u16, a1: &str, formula: &str) -> &mut Self {
        let at = Self::at(sheet, a1);
        let formula = formula.strip_prefix('=').unwrap_or(formula).to_string();
        self.cells.insert(
            at,
            ModelCell {
                value: Value::Empty,
                formula: Some(formula),
                array: false,
            },
        );
        self
    }

    #[must_use]
    pub fn sheet_id(&self, name: &str) -> Option<u16> {
        self.sheets
            .iter()
            .position(|s| sheetrs_engine::addr::same_name(s, name))
            .map(|i| i as u16)
    }
}

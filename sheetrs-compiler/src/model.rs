//! In-memory, format-independent workbook: what every frontend produces and
//! the compiler consumes. Formulas stay as text until compilation, once all
//! sheet and defined names are known.

use crate::addr::{self, CellRef};
use crate::cellstore::CellStore;
use crate::options::{DateSystem, Iteration, Search};
use crate::value::Value;
use std::collections::HashMap;

#[derive(Debug, Clone, PartialEq)]
pub struct ModelCell {
    /// Constant value, or the cached result of the formula.
    pub value: Value,
    /// Formula text without the leading `=`, in Excel syntax.
    pub formula: Option<String>,
    /// Array (Ctrl+Shift+Enter) formula: its operators work element-wise.
    pub array: bool,
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
    /// (Calc A1, a document created in LibreOffice). `None`: Excel's.
    pub string_ref_excel: Option<bool>,
    /// The document is in manual calculation mode (it calculates only on request).
    pub manual_calc: bool,
    /// Document language tag (`es-CO`).
    pub locale: Option<String>,
}

#[derive(Debug, Clone)]
pub struct Model {
    pub sheets: Vec<String>,
    pub cells: CellStore,
    /// Defined names: name → formula text (usually a reference).
    pub names: HashMap<String, String>,
    pub date_system: DateSystem,
    pub settings: ModelSettings,
    /// Built by a file loader: the document defines date system, iteration,
    /// search settings and reference syntax, and a `Spec` may not override them.
    pub from_file: bool,
    /// The formulas come from an XLSX (LibreOffice reads a bare `#REF!` in one as an
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
    pub fn new() -> Self {
        Self {
            sheets: Vec::new(),
            cells: CellStore::default(),
            names: HashMap::new(),
            // LibreOffice numbers dates from 1899-12-30 on the real calendar.
            date_system: DateSystem::NULL_1899,
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

    fn at(&self, sheet: u16, a1: &str) -> CellRef {
        let area = addr::parse_area(a1).unwrap_or_else(|| panic!("bad address `{a1}`"));
        CellRef::new(sheet, area.r0, area.c0)
    }

    /// Set a constant, addressing the cell in A1 notation.
    pub fn set(&mut self, sheet: u16, a1: &str, value: Value) -> &mut Self {
        let at = self.at(sheet, a1);
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
        let at = self.at(sheet, a1);
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

    pub fn sheet_id(&self, name: &str) -> Option<u16> {
        self.sheets
            .iter()
            .position(|s| crate::addr::same_name(s, name))
            .map(|i| i as u16)
    }
}

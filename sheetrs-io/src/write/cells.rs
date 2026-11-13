//! Cells written with values someone else decided: what a program computed, or what an input is.
//!
//! [`CellWrites`] is the other half of `formula_to_value`: that one takes the results a workbook
//! caches; this one puts results the caller has (the answer of a compiled program, say) in a
//! workbook that serves as a template. The formulas, the arrays and the styles stay as they are, the
//! cached results become the new ones, and the constants named are replaced, so the file that comes
//! out is the template with the numbers of this run, ready to be opened.

use crate::write::{CachedResult, CellEdit, SheetEdits};
use std::collections::HashMap;

/// The constant a cell holds.
#[derive(Debug, Clone, PartialEq)]
pub enum Constant {
    Number(f64),
    Text(String),
    Boolean(bool),
    /// The cell is emptied.
    Empty,
}

/// Values to write into the cells of a workbook, by sheet name and `(row, column)`, 0-based.
///
/// The cached results written are taken as current: the saved workbook is not marked to be
/// recalculated when it is opened because of them.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CellWrites {
    sheets: HashMap<String, SheetEdits>,
}

impl CellWrites {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// The cell keeps what it is (its formula, its array formula, its style) and shows `result`.
    pub fn result(
        &mut self,
        sheet: impl Into<String>,
        row: u32,
        col: u32,
        result: CachedResult,
    ) -> &mut Self {
        self.sheets
            .entry(sheet.into())
            .or_default()
            .insert((row, col), CellEdit::Cached(result));
        self
    }

    /// The cell holds `constant` (a formula it had is gone), keeping its style.
    pub fn constant(
        &mut self,
        sheet: impl Into<String>,
        row: u32,
        col: u32,
        constant: Constant,
    ) -> &mut Self {
        let edit = match constant {
            Constant::Number(n) => CellEdit::Number(n),
            Constant::Text(t) => CellEdit::Text(t),
            Constant::Boolean(b) => CellEdit::Boolean(b),
            Constant::Empty => CellEdit::Clear,
        };
        self.sheets
            .entry(sheet.into())
            .or_default()
            .insert((row, col), edit);
        self
    }

    /// The cells written, by sheet name: what a writer is given as the replacements of a pass.
    #[must_use]
    pub fn sheets(&self) -> &HashMap<String, SheetEdits> {
        &self.sheets
    }

    /// How many cells are written.
    #[must_use]
    pub fn len(&self) -> usize {
        self.sheets.values().map(SheetEdits::len).sum()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }
}

//! VUL601: Duplicate formulas
//!
//! The same formula text in several cells of a sheet points at copy and paste, at a calculation that
//! could be kept once (a helper cell, a named range or a dynamic array), and at a risk of the copies
//! being corrected separately.
//!
//! # Reports
//! One violation per group of adjacent cells (scope: sheet) holding the same formula. Within a sheet,
//! formulas are compared as text after trimming surrounding white space; a text that occurs in
//! at least two cells forms a candidate, and its cells are split into groups of cells that touch
//! horizontally or vertically (a contiguous block). Each group gives one violation with its bounding
//! box. The message shows the formula (cut at 50 characters), the total number of cells with
//! that text on the sheet (the same number for each of its groups) and the group's range. A formula
//! that occurs once is never reported. Formulas of any kind and length are compared, including
//! trivial ones.
//!
//! # Does not report
//! Formulas that differ in text, even when they compute the same thing (`=A1+B1` and `=B1+A1`), and
//! the same formula on different sheets: the comparison is per sheet. Fill-down formulas whose
//! references change from row to row are different text and are not duplicates.
//!
//! # Parameters
//! None.
//!
//! # Severity
//! Always `Info`: a duplicate is a hint, not necessarily a fault.

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use sheetrs_io::reader::{Cell, Sheet};
use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};

/// Rule that detects duplicate formulas in non-adjacent cells
#[derive(Debug)]
pub struct DuplicateFormulasRule;

impl DuplicateFormulasRule {
    /// Create a new instance.
    #[must_use]
    pub fn new() -> Self {
        Self
    }
}

impl Default for DuplicateFormulasRule {
    fn default() -> Self {
        Self::new()
    }
}

/// Incident data for VUL601.
#[derive(Debug)]
pub struct DuplicateFormulaData {
    /// The duplicated formula text.
    pub formula: String,
    /// Number of times the formula appears.
    pub count: usize,
    /// Bounding box of one contiguous group (`start_row`, `start_col`, `end_row`, `end_col`).
    pub range: (u32, u32, u32, u32),
}

impl ViolationData for DuplicateFormulaData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let (sr, sc, er, ec) = self.range;
        let range_str = if sr == er && sc == ec {
            CellReference::new(sr, sc).to_string()
        } else {
            format!(
                "{}:{}",
                CellReference::new(sr, sc),
                CellReference::new(er, ec)
            )
        };

        format!(
            "Formula '{}' is duplicated {} times in ranges: {}. Consider using named ranges or helper cells.",
            self.formula, self.count, range_str
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for DuplicateFormulasRule {
    fn id(&self) -> RuleId {
        RuleId::Vul601
    }

    fn name(&self) -> &'static str {
        "Duplicate Formula"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Vulnerability
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext, out: &mut Vec<Violation>) {
        // A map of every formula text held a hash entry for each (a sheet of millions of formulas
        // that are nearly all different cost ~250 MB). Each formula is kept as the hash of its
        // text and its position, 16 bytes; sorted, the equal hashes are adjacent, and only
        // those runs are compared as text (a collision cannot join different formulas).
        let mut hashed: Vec<(u64, u32, u32)> = Vec::with_capacity(sheet.cells.len());
        for ((row, col), cell) in sheet.cells.iter() {
            if let Some(formula) = cell.as_formula() {
                let mut hasher = DefaultHasher::new();
                formula.trim().hash(&mut hasher);
                hashed.push((hasher.finish(), row, col));
            }
        }
        hashed.sort_unstable();

        // The formulas are the sheet's own text, borrowed: nothing is copied but the few that
        // are reported.
        let mut duplicates: Vec<(&str, Vec<(u32, u32)>)> = Vec::new();
        for run in hashed.chunk_by(|a, b| a.0 == b.0) {
            if run.len() < 2 {
                continue;
            }
            let mut texts: Vec<(&str, Vec<(u32, u32)>)> = Vec::new();
            for &(_, row, col) in run {
                let Some(formula) = sheet.cell(row, col).and_then(Cell::as_formula) else {
                    continue;
                };
                let formula = formula.trim();
                match texts.iter_mut().find(|(text, _)| *text == formula) {
                    Some((_, cells)) => cells.push((row, col)),
                    None => texts.push((formula, vec![(row, col)])),
                }
            }
            duplicates.extend(texts.into_iter().filter(|(_, cells)| cells.len() > 1));
        }
        drop(hashed);
        duplicates.sort_unstable_by_key(|(formula, _)| *formula);

        for (formula, mut cells) in duplicates {
            cells.sort_unstable();
            let ranges = find_contiguous_ranges(&cells);

            // Truncate formula for display
            let display_formula = if formula.chars().count() > 50 {
                formula.chars().take(50).collect::<String>() + "..."
            } else {
                formula.to_string()
            };

            let count = cells.len();

            for range in ranges {
                let bbox = bounding_box(&range);
                out.push(Violation::with_data(
                    RuleId::Vul601,
                    ViolationScope::Sheet(sheet.sheet_index),
                    DuplicateFormulaData {
                        formula: display_formula.clone(),
                        count,
                        range: bbox,
                    },
                    Severity::Info,
                ));
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::rules::collect_violations;
    use sheetrs_io::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;

    #[test]
    fn test_duplicate_formulas() {
        let mut cells = HashMap::new();
        for row in 0..3 {
            cells.insert(
                (row, 0),
                Cell {
                    formula: Some(<Box<str>>::from("=A1+B1")),
                    is_array: false,
                    num_fmt: None,
                    value: CellValue::Empty,
                },
            );
        }

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            ..Default::default()
        };

        let mut rule = DuplicateFormulasRule::new();
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Vul601);
    }

    #[test]
    fn test_unique_formulas() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=A1+B1")),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );
        cells.insert(
            (1, 0),
            Cell {
                formula: Some(<Box<str>>::from("=A2+B2")),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            ..Default::default()
        };

        let mut rule = DuplicateFormulasRule::new();
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert!(violations.is_empty());
    }
}

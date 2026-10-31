//! VUL605: Legacy array formulas
//!
//! Array formulas entered with Ctrl+Shift+Enter (CSE) are fixed to a range of cells and are easy to
//! edit wrongly: changing one cell of the array is refused, and the braces hide that a formula is an
//! array. Dynamic array formulas spill by themselves.
//!
//! # Reports
//! One violation per group of adjacent cells (scope: sheet) that hold an array formula; the message
//! gives the group's bounding box. A cell is included when the reader marks it as an array formula:
//! an XLSX formula with `t="array"`, or an ODS matrix formula, and it has a formula. No other
//! property of the formula is checked, so the rule does not tell a formula that needs CSE from one
//! that would work unchanged as a normal formula, nor a CSE array from a dynamic array stored
//! with the same marker.
//!
//! # Does not report
//! Ordinary formulas, including those that use array arithmetic without being stored as array
//! formulas, and cells that carry the array flag without a formula.
//!
//! # Parameters
//! None.
//!
//! # Severity
//! Always `Warning`.

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::HashMap;

/// Rule that identifies legacy CSE array formulas.
pub struct LegacyArrayRule {
    /// Per-sheet collected cells: sheet_index → Vec<(row, col)>.
    sheet_cells: HashMap<u16, Vec<(u32, u32)>>,
}

impl LegacyArrayRule {
    pub fn new() -> Self {
        Self {
            sheet_cells: HashMap::new(),
        }
    }
}

impl Default for LegacyArrayRule {
    fn default() -> Self {
        Self::new()
    }
}

/// Incident data for VUL605.
#[derive(Debug)]
pub struct LegacyArrayData {
    /// Bounding box of the CSE array range (start_row, start_col, end_row, end_col).
    pub range: (u32, u32, u32, u32),
}

impl ViolationData for LegacyArrayData {
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

        format!("Legacy array formula (CSE) found in range: {range_str}. Convert to dynamic array.")
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for LegacyArrayRule {
    fn id(&self) -> RuleId {
        RuleId::Vul605
    }

    fn name(&self) -> &'static str {
        "Legacy Array"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Vulnerability
    }

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        at: (u32, u32),
        cell: &Cell,
        _ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        if cell.is_array && cell.is_formula() {
            self.sheet_cells
                .entry(sheet.sheet_index)
                .or_default()
                .push((at.0, at.1));
        }
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext, out: &mut Vec<Violation>) {
        let Some(mut cells) = self.sheet_cells.remove(&sheet.sheet_index) else {
            return;
        };
        cells.sort_unstable();
        let ranges = find_contiguous_ranges(&cells);

        for range in ranges {
            let bbox = bounding_box(&range);
            out.push(Violation::with_data(
                RuleId::Vul605,
                ViolationScope::Sheet(sheet.sheet_index),
                LegacyArrayData { range: bbox },
                Severity::Warning,
            ));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::LinterContext;
    use crate::rules::collect_violations;
    use std::collections::HashMap;

    fn make_cell(row: u32, col: u32, formula: Option<&str>, is_array: bool) -> ((u32, u32), Cell) {
        (
            (row, col),
            Cell {
                formula: formula.map(Box::from),
                num_fmt: None,
                value: CellValue::Empty,
                is_array,
            },
        )
    }

    fn make_sheet(cells: Vec<((u32, u32), Cell)>) -> Sheet {
        let mut cell_map = HashMap::new();
        for (at, cell) in cells {
            cell_map.insert(at, cell);
        }
        Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cell_map.into(),
            ..Default::default()
        }
    }

    fn run_rule(rule: &mut LegacyArrayRule, sheet: &Sheet) -> Vec<Violation> {
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        collect_violations(|out| rule.on_sheet_end(sheet, &mut ctx, out))
    }

    #[test]
    fn test_single_cse_cell() {
        let sheet = make_sheet(vec![make_cell(0, 0, Some("SUM(A2:A10)"), true)]);
        let mut rule = LegacyArrayRule::new();
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Vul605);
    }

    #[test]
    fn test_non_cse_formula() {
        let sheet = make_sheet(vec![make_cell(0, 0, Some("SUM(A2:A10)"), false)]);
        let mut rule = LegacyArrayRule::new();
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_multi_cell_cse_block_grouped() {
        let sheet = make_sheet(vec![
            make_cell(0, 0, Some("SUM(A2:A10)"), true),
            make_cell(0, 1, Some("SUM(A2:A10)"), true),
            make_cell(1, 0, Some("SUM(A2:A10)"), true),
        ]);
        let mut rule = LegacyArrayRule::new();
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 1, "Adjacent CSE cells should be grouped");
    }

    #[test]
    fn test_no_formula_with_is_array() {
        // Value-only cell with bogus is_array flag → no violation
        let sheet = make_sheet(vec![{
            let (at, mut c) = make_cell(0, 0, None, true);
            c.value = CellValue::Number(42.0);
            (at, c)
        }]);
        let mut rule = LegacyArrayRule::new();
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_separate_cse_ranges() {
        let sheet = make_sheet(vec![
            make_cell(0, 0, Some("SUM(A2:A10)"), true),
            make_cell(5, 5, Some("SUM(B2:B10)"), true),
        ]);
        let mut rule = LegacyArrayRule::new();
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(
            violations.len(),
            2,
            "Non-adjacent CSE cells should produce separate violations"
        );
    }
}

//! VUL605: Legacy Array detection
//!
//! Description: Detects antiquated CSE (Ctrl+Shift+Enter) array formulas
//! that should be converted to modern dynamic array spill formulas.

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::HashMap;
use std::sync::Mutex;

/// Rule that identifies legacy CSE array formulas.
pub struct LegacyArrayRule {
    /// Per-sheet collected cells: sheet_index → Vec<(row, col)>.
    sheet_cells: Mutex<HashMap<u16, Vec<(u32, u32)>>>,
}

impl LegacyArrayRule {
    pub fn new() -> Self {
        Self {
            sheet_cells: Mutex::new(HashMap::new()),
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

    fn on_cell(&self, sheet: &Sheet, cell: &Cell, _ctx: &mut LinterContext) -> Vec<Violation> {
        if cell.is_array && cell.is_formula() {
            let mut map = self.sheet_cells.lock().unwrap();
            map.entry(sheet.sheet_index)
                .or_default()
                .push((cell.row, cell.col));
        }
        Vec::new()
    }

    fn on_sheet_end(&self, sheet: &Sheet, _ctx: &mut LinterContext) -> Vec<Violation> {
        let Some(cells) = self.sheet_cells.lock().unwrap().remove(&sheet.sheet_index) else {
            return Vec::new();
        };

        let mut violations = Vec::new();
        let ranges = find_contiguous_ranges(&cells);

        for range in ranges {
            let bbox = bounding_box(&range);
            violations.push(Violation::with_data(
                RuleId::Vul605,
                ViolationScope::Sheet(sheet.sheet_index),
                LegacyArrayData { range: bbox },
                Severity::Warning,
            ));
        }

        violations
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::LinterContext;
    use std::collections::HashMap;

    fn make_cell(row: u32, col: u32, formula: Option<&str>, is_array: bool) -> Cell {
        Cell {
            formula: formula.map(Box::from),
            num_fmt: None,
            row,
            col,
            value: CellValue::Empty,
            is_array,
        }
    }

    fn make_sheet(cells: Vec<Cell>) -> Sheet {
        let mut cell_map = HashMap::new();
        for cell in cells {
            cell_map.insert((cell.row, cell.col), cell);
        }
        Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cell_map,
            ..Default::default()
        }
    }

    fn run_rule(rule: &LegacyArrayRule, sheet: &Sheet) -> Vec<Violation> {
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(sheet, cell, &mut ctx);
        }
        rule.on_sheet_end(sheet, &mut ctx)
    }

    #[test]
    fn test_single_cse_cell() {
        let sheet = make_sheet(vec![make_cell(0, 0, Some("SUM(A2:A10)"), true)]);
        let rule = LegacyArrayRule::new();
        let violations = run_rule(&rule, &sheet);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Vul605);
    }

    #[test]
    fn test_non_cse_formula() {
        let sheet = make_sheet(vec![make_cell(0, 0, Some("SUM(A2:A10)"), false)]);
        let rule = LegacyArrayRule::new();
        let violations = run_rule(&rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_multi_cell_cse_block_grouped() {
        let sheet = make_sheet(vec![
            make_cell(0, 0, Some("SUM(A2:A10)"), true),
            make_cell(0, 1, Some("SUM(A2:A10)"), true),
            make_cell(1, 0, Some("SUM(A2:A10)"), true),
        ]);
        let rule = LegacyArrayRule::new();
        let violations = run_rule(&rule, &sheet);

        assert_eq!(violations.len(), 1, "Adjacent CSE cells should be grouped");
    }

    #[test]
    fn test_no_formula_with_is_array() {
        // Value-only cell with bogus is_array flag → no violation
        let sheet = make_sheet(vec![{
            let mut c = make_cell(0, 0, None, true);
            c.value = CellValue::Number(42.0);
            c
        }]);
        let rule = LegacyArrayRule::new();
        let violations = run_rule(&rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_separate_cse_ranges() {
        let sheet = make_sheet(vec![
            make_cell(0, 0, Some("SUM(A2:A10)"), true),
            make_cell(5, 5, Some("SUM(B2:B10)"), true),
        ]);
        let rule = LegacyArrayRule::new();
        let violations = run_rule(&rule, &sheet);

        assert_eq!(
            violations.len(),
            2,
            "Non-adjacent CSE cells should produce separate violations"
        );
    }
}

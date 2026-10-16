//! VUL606: Deprecated Func detection
//!
//! Description: Identifies functions superseded by modern alternatives (e.g., CONCATENATE vs CONCAT).

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::{BTreeMap, HashMap};
use std::sync::LazyLock;

/// Deprecated function entries: (deprecated_name, replacement_name).
const DEPRECATED_FUNCTIONS: &[(&str, &str)] = &[
    ("CONCATENATE", "CONCAT / TEXTJOIN"), // Excel 2019
    ("LOOKUP", "XLOOKUP / INDEX+MATCH"),  // Excel 365
    ("FORECAST", "FORECAST.LINEAR"),      // Excel 2016
    ("STDEV", "STDEV.S"),                 // Excel 2010
    ("STDEVP", "STDEV.P"),                // Excel 2010
    ("VAR", "VAR.S"),                     // Excel 2010
    ("VARP", "VAR.P"),                    // Excel 2010
    ("COVAR", "COVARIANCE.P"),            // Excel 2010
    ("MODE", "MODE.SNGL"),                // Excel 2010
    ("PERCENTILE", "PERCENTILE.INC"),     // Excel 2010
    ("PERCENTRANK", "PERCENTRANK.INC"),   // Excel 2010
    ("QUARTILE", "QUARTILE.INC"),         // Excel 2010
    ("RANK", "RANK.EQ"),                  // Excel 2010
    ("CEILING", "CEILING.MATH"),          // Excel 2013
    ("FLOOR", "FLOOR.MATH"),              // Excel 2013
];

/// Pre-computed search patterns: "FUNC(" for each deprecated function.
static SEARCH_PATTERNS: LazyLock<Vec<String>> = LazyLock::new(|| {
    DEPRECATED_FUNCTIONS
        .iter()
        .map(|(name, _)| format!("{name}("))
        .collect()
});

/// Per-sheet cell collection: sheet_index → Vec<(func_index, row, col)>.
type SheetCellMap = HashMap<u16, Vec<(u8, u32, u32)>>;

/// Rule that detects usage of deprecated functions.
pub struct DeprecatedFuncRule {
    /// Per-sheet collected cells with function index.
    sheet_cells: SheetCellMap,
}

impl DeprecatedFuncRule {
    /// Create a new instance.
    pub fn new() -> Self {
        Self {
            sheet_cells: HashMap::new(),
        }
    }
}

impl Default for DeprecatedFuncRule {
    fn default() -> Self {
        Self::new()
    }
}

/// Incident data for VUL606.
#[derive(Debug)]
pub struct DeprecatedFuncData {
    /// Index into `DEPRECATED_FUNCTIONS` list.
    pub func_index: u8,
    /// Bounding box of violating cells (start_row, start_col, end_row, end_col).
    pub range: (u32, u32, u32, u32),
}

impl ViolationData for DeprecatedFuncData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let (deprecated, replacement) = DEPRECATED_FUNCTIONS
            .get(self.func_index as usize)
            .unwrap_or(&("UNKNOWN", "UNKNOWN"));

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
            "Deprecated function {deprecated}() found in range: {range_str}. Use {replacement}() instead."
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for DeprecatedFuncRule {
    fn id(&self) -> RuleId {
        RuleId::Vul606
    }

    fn name(&self) -> &'static str {
        "Deprecated Func"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Vulnerability
    }

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        cell: &Cell,
        ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        if let Some(formula_upper) = ctx.formula_upper(sheet, cell) {
            let formula_bytes = formula_upper.as_bytes();
            let mut matches = Vec::new();
            for (idx, pattern) in SEARCH_PATTERNS.iter().enumerate() {
                // Find all occurrences and check word boundary before each
                let pat_bytes = pattern.as_bytes();
                let mut search_from = 0;
                while let Some(rel_pos) = formula_upper[search_from..].find(pattern.as_str()) {
                    let abs_pos = search_from + rel_pos;
                    // Only match if NOT preceded by a letter (rejects VLOOKUP→LOOKUP, etc.)
                    let is_word_start =
                        abs_pos == 0 || !formula_bytes[abs_pos - 1].is_ascii_alphabetic();
                    if is_word_start {
                        matches.push((idx as u8, cell.row, cell.col));
                        break; // Found a genuine match for this pattern in this cell
                    }
                    search_from = abs_pos + pat_bytes.len();
                }
            }
            if !matches.is_empty() {
                self.sheet_cells
                    .entry(sheet.sheet_index)
                    .or_default()
                    .extend(matches);
            }
        }
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext, out: &mut Vec<Violation>) {
        let Some(mut cells) = self.sheet_cells.remove(&sheet.sheet_index) else {
            return;
        };
        cells.sort_unstable();

        // Group by function index
        let mut by_func: BTreeMap<u8, Vec<(u32, u32)>> = BTreeMap::new();
        for (func_idx, row, col) in cells {
            by_func.entry(func_idx).or_default().push((row, col));
        }

        for (func_idx, func_cells) in by_func {
            let ranges = find_contiguous_ranges(&func_cells);
            for range in ranges {
                let bbox = bounding_box(&range);
                out.push(Violation::with_data(
                    RuleId::Vul606,
                    ViolationScope::Sheet(sheet.sheet_index),
                    DeprecatedFuncData {
                        func_index: func_idx,
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
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::collect_violations;
    use std::collections::HashMap;

    fn make_formula_cell(row: u32, col: u32, formula: &str) -> Cell {
        Cell {
            formula: Some(<Box<str>>::from(formula)),
            is_array: false,
            num_fmt: None,
            row,
            col,
            value: CellValue::Empty,
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

    fn run_rule(rule: &mut DeprecatedFuncRule, sheet: &Sheet) -> Vec<Violation> {
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(sheet, cell, &mut ctx, &mut Vec::new());
        }
        collect_violations(|out| rule.on_sheet_end(sheet, &mut ctx, out))
    }

    #[test]
    fn test_single_deprecated_function() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "CONCATENATE(A2,B2)")]);
        let mut rule = DeprecatedFuncRule::new();
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Vul606);
    }

    #[test]
    fn test_replacement_not_flagged() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "CONCAT(A2,B2)")]);
        let mut rule = DeprecatedFuncRule::new();
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_case_insensitive() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "concatenate(A2,B2)")]);
        let mut rule = DeprecatedFuncRule::new();
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_multiple_deprecated() {
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "STDEV(A1:A10)"),
            make_formula_cell(1, 0, "RANK(B1,B1:B10)"),
        ]);
        let mut rule = DeprecatedFuncRule::new();
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 2);
    }

    #[test]
    fn test_substring_not_false_positive() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "STDEV.S(A1:A10)")]);
        let mut rule = DeprecatedFuncRule::new();
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_no_formula() {
        let mut cell_map = HashMap::new();
        cell_map.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Number(42.0),
            },
        );
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cell_map,
            ..Default::default()
        };
        let mut rule = DeprecatedFuncRule::new();
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_vlookup_not_false_positive() {
        // VLOOKUP contains LOOKUP but should NOT trigger the deprecated LOOKUP rule
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "VLOOKUP(A1,B:C,2,FALSE)")]);
        let mut rule = DeprecatedFuncRule::new();
        let violations = run_rule(&mut rule, &sheet);
        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_hlookup_not_false_positive() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "HLOOKUP(A1,1:3,2,FALSE)")]);
        let mut rule = DeprecatedFuncRule::new();
        let violations = run_rule(&mut rule, &sheet);
        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_xlookup_not_false_positive() {
        // XLOOKUP is the recommended replacement for LOOKUP — must NOT be flagged
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "XLOOKUP(A1,B1:B10,C1:C10)")]);
        let mut rule = DeprecatedFuncRule::new();
        let violations = run_rule(&mut rule, &sheet);
        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_bare_lookup_still_flagged() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "LOOKUP(A1,B1:B10,C1:C10)")]);
        let mut rule = DeprecatedFuncRule::new();
        let violations = run_rule(&mut rule, &sheet);
        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_covar_not_flagged_as_var() {
        // COVAR contains VAR but should only flag COVAR, not VAR
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "COVAR(A1:A10,B1:B10)")]);
        let mut rule = DeprecatedFuncRule::new();
        let violations = run_rule(&mut rule, &sheet);
        // Should flag COVAR only, not also VAR
        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_two_deprecated_in_one_cell() {
        // A single cell with two different deprecated functions should produce two violations
        let sheet = make_sheet(vec![make_formula_cell(
            0,
            0,
            "CONCATENATE(STDEV(A1:A10),B1)",
        )]);
        let mut rule = DeprecatedFuncRule::new();
        let violations = run_rule(&mut rule, &sheet);
        assert_eq!(violations.len(), 2);
    }
}

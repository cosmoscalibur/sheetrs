//! VUL604: Error Prone Functions detection
//!
//! Description: Detects Excel functions that are susceptible to generating
//! errors or returning unexpected results. Flags VLOOKUP, HLOOKUP, and LOOKUP
//! as error-prone; recommends XLOOKUP or INDEX/MATCH alternatives.

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::HashMap;
use std::sync::{LazyLock, Mutex};

/// Error-prone function entries: (function_name, recommended_alternative).
const ERROR_PRONE_FUNCTIONS: &[(&str, &str)] = &[
    ("VLOOKUP", "XLOOKUP / INDEX+MATCH"),
    ("HLOOKUP", "XLOOKUP / INDEX+MATCH"),
    ("LOOKUP", "XLOOKUP / INDEX+MATCH"),
];

/// Pre-computed search patterns: "FUNC(" for each error-prone function.
static SEARCH_PATTERNS: LazyLock<Vec<String>> = LazyLock::new(|| {
    ERROR_PRONE_FUNCTIONS
        .iter()
        .map(|(name, _)| format!("{name}("))
        .collect()
});

/// Per-sheet cell collection: sheet_index → Vec<(func_index, row, col)>.
type SheetCellMap = Mutex<HashMap<u16, Vec<(u8, u32, u32)>>>;

/// Rule that detects error-prone functions like VLOOKUP, HLOOKUP, and LOOKUP.
pub struct ErrorProneFunctionsRule {
    /// Per-sheet collected cells with function index.
    sheet_cells: SheetCellMap,
}

impl ErrorProneFunctionsRule {
    pub fn new(_config: &LinterConfig) -> Self {
        Self {
            sheet_cells: Mutex::new(HashMap::new()),
        }
    }
}

/// Incident data for VUL604.
#[derive(Debug)]
pub struct ErrorProneFuncData {
    /// Index into `ERROR_PRONE_FUNCTIONS` list.
    pub func_index: u8,
    /// Bounding box of violating cells (start_row, start_col, end_row, end_col).
    pub range: (u32, u32, u32, u32),
}

impl ViolationData for ErrorProneFuncData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let (func_name, alternative) = ERROR_PRONE_FUNCTIONS
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
            "Error-prone function {func_name}() found in range: {range_str}. Use {alternative} instead."
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for ErrorProneFunctionsRule {
    fn id(&self) -> RuleId {
        RuleId::Vul604
    }

    fn name(&self) -> &'static str {
        "Error Prone Functions"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Vulnerability
    }

    fn on_cell(&self, sheet: &Sheet, cell: &Cell, ctx: &mut LinterContext) -> Vec<Violation> {
        if let Some(formula_upper) = ctx.formula_upper(sheet, cell) {
            let formula_bytes = formula_upper.as_bytes();
            let mut matches = Vec::new();

            for (idx, pattern) in SEARCH_PATTERNS.iter().enumerate() {
                let pat_bytes = pattern.as_bytes();
                let mut search_from = 0;

                while let Some(rel_pos) = formula_upper[search_from..].find(pattern.as_str()) {
                    let abs_pos = search_from + rel_pos;

                    // Word boundary: reject XLOOKUP→LOOKUP, VLOOKUP→LOOKUP, etc.
                    let is_word_start =
                        abs_pos == 0 || !formula_bytes[abs_pos - 1].is_ascii_alphabetic();

                    if is_word_start {
                        matches.push((idx as u8, cell.row, cell.col));
                        break; // One match per function per cell
                    }

                    search_from = abs_pos + pat_bytes.len();
                }
            }
            if !matches.is_empty() {
                self.sheet_cells
                    .lock()
                    .unwrap()
                    .entry(sheet.sheet_index)
                    .or_default()
                    .extend(matches);
            }
        }
        Vec::new()
    }

    fn on_sheet_end(&self, sheet: &Sheet, _ctx: &mut LinterContext) -> Vec<Violation> {
        let Some(cells) = self.sheet_cells.lock().unwrap().remove(&sheet.sheet_index) else {
            return Vec::new();
        };

        let mut violations = Vec::new();

        // Group by function index
        let mut by_func: HashMap<u8, Vec<(u32, u32)>> = HashMap::new();
        for (func_idx, row, col) in cells {
            by_func.entry(func_idx).or_default().push((row, col));
        }

        for (func_idx, func_cells) in by_func {
            let ranges = find_contiguous_ranges(&func_cells);
            for range in ranges {
                let bbox = bounding_box(&range);
                violations.push(Violation::with_data(
                    RuleId::Vul604,
                    ViolationScope::Sheet(sheet.sheet_index),
                    ErrorProneFuncData {
                        func_index: func_idx,
                        range: bbox,
                    },
                    Severity::Warning,
                ));
            }
        }

        violations
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
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

    fn run_rule(rule: &ErrorProneFunctionsRule, sheet: &Sheet) -> Vec<Violation> {
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(sheet, cell, &mut ctx);
        }
        rule.on_sheet_end(sheet, &mut ctx)
    }

    #[test]
    fn test_vlookup_detection() {
        let config = LinterConfig::default();
        let rule = ErrorProneFunctionsRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "VLOOKUP(A1,B:C,2,FALSE)")]);
        let violations = run_rule(&rule, &sheet);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Vul604);
    }

    #[test]
    fn test_hlookup_detection() {
        let config = LinterConfig::default();
        let rule = ErrorProneFunctionsRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "HLOOKUP(A1,1:3,2,FALSE)")]);
        let violations = run_rule(&rule, &sheet);

        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_lookup_detection() {
        let config = LinterConfig::default();
        let rule = ErrorProneFunctionsRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "LOOKUP(A1,B1:B10,C1:C10)")]);
        let violations = run_rule(&rule, &sheet);

        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_xlookup_no_false_positive() {
        let config = LinterConfig::default();
        let rule = ErrorProneFunctionsRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "XLOOKUP(A1,B1:B10,C1:C10)")]);
        let violations = run_rule(&rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_sum_not_flagged() {
        let config = LinterConfig::default();
        let rule = ErrorProneFunctionsRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "SUM(A1:A10)")]);
        let violations = run_rule(&rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_case_insensitive() {
        let config = LinterConfig::default();
        let rule = ErrorProneFunctionsRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "vlookup(A1,B:C,2,FALSE)")]);
        let violations = run_rule(&rule, &sheet);

        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_multiple_functions_separate_violations() {
        let config = LinterConfig::default();
        let rule = ErrorProneFunctionsRule::new(&config);
        let sheet = make_sheet(vec![
            make_formula_cell(0, 0, "VLOOKUP(A1,B:C,2,FALSE)"),
            make_formula_cell(0, 1, "HLOOKUP(A1,B:C,2,FALSE)"),
        ]);
        let violations = run_rule(&rule, &sheet);

        // VLOOKUP and HLOOKUP are different func_index → separate groups
        assert_eq!(violations.len(), 2);
    }

    #[test]
    fn test_no_formula() {
        let config = LinterConfig::default();
        let rule = ErrorProneFunctionsRule::new(&config);
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
        let violations = run_rule(&rule, &sheet);

        assert_eq!(violations.len(), 0);
    }
}

//! CALC204: Approximate Lookup
//!
//! `VLOOKUP`, `HLOOKUP`, `MATCH` and `LOOKUP` search for the closest value when no exact-match flag
//! is given, and only work if the data is sorted. On unsorted data they return a wrong row without
//! any error, so the result looks valid and is not.
//!
//! # Reports
//!
//! One violation per function per formula cell (unit: the cell and function; a cell using both
//! `VLOOKUP` and `MATCH` unmatched gets two). A call is reported when:
//!
//! - `VLOOKUP` or `HLOOKUP`: the 4th argument is missing or anything other than the literal `FALSE`
//!   or `0` (so `TRUE`, `1`, a cell reference or a nested expression is reported).
//! - `MATCH`: the 3rd argument is missing or anything other than the literal `0` (`1` and `-1` are
//!   approximate modes and are reported).
//! - `LOOKUP`: always, since it has no exact-match flag. The message points to `XLOOKUP` or
//!   `INDEX`/`MATCH`.
//!
//! If one call of a function is exact and another is not, the unmatched call is still reported.
//! Function names are matched case-insensitively; a malformed argument list counts as not exact.
//!
//! # Does not report
//!
//! - `XLOOKUP`, which matches exactly by default; the name must start at a word boundary, so it is
//!   not mistaken for `LOOKUP`.
//! - Function names inside string constants.
//!
//! The rule cannot tell deliberate approximate matching (banded rates, tax tables) from a mistake,
//! so such calls are reported too; stating the mode explicitly does not silence it, because only
//! the literals `FALSE` and `0` count as exact.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always `Warning`.

use super::helpers::{extract_args, is_inside_string};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::sync::LazyLock;

/// Pre-computed search patterns: "VLOOKUP(", "HLOOKUP(", "MATCH(", "LOOKUP(".
static SEARCH_PATTERNS: LazyLock<Vec<String>> = LazyLock::new(|| {
    LOOKUP_FUNCTIONS
        .iter()
        .map(|(name, _)| format!("{name}("))
        .collect()
});

/// Rule that identifies approximate lookup functions.
pub struct ApproximateLookupRule;

/// Incident data for CALC204.
#[derive(Debug)]
pub struct ApproximateLookupData {
    /// The function name that uses approximate match.
    pub function: &'static str,
}

impl ViolationData for ApproximateLookupData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        match self.function {
            "VLOOKUP" => "VLOOKUP uses approximate match (default). Add FALSE as 4th argument for exact match.".to_string(),
            "HLOOKUP" => "HLOOKUP uses approximate match (default). Add FALSE as 4th argument for exact match.".to_string(),
            "MATCH" => "MATCH uses approximate match (default). Set 3rd argument to 0 for exact match.".to_string(),
            "LOOKUP" => "LOOKUP always uses approximate match. Use XLOOKUP or INDEX/MATCH for exact match.".to_string(),
            _ => format!("{} uses approximate match.", self.function),
        }
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Functions to check and the arg index (0-based) of the exact-match flag.
/// `None` means the function always uses approximate match (no flag).
const LOOKUP_FUNCTIONS: &[(&str, Option<usize>)] = &[
    ("VLOOKUP", Some(3)), // 4th arg: FALSE/0 = exact
    ("HLOOKUP", Some(3)), // 4th arg: FALSE/0 = exact
    ("MATCH", Some(2)),   // 3rd arg: 0 = exact
    ("LOOKUP", None),     // Always approximate, no flag
];

impl WalkerRule for ApproximateLookupRule {
    fn id(&self) -> RuleId {
        RuleId::Calc204
    }

    fn name(&self) -> &'static str {
        "Approximate Lookup"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Calculations
    }

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        at: (u32, u32),
        cell: &Cell,
        ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        let Some(formula_upper) = ctx.formula_upper(sheet, at, cell) else {
            return;
        };
        let formula_bytes = formula_upper.as_bytes();

        for (i, &(func_name, exact_match_arg_idx)) in LOOKUP_FUNCTIONS.iter().enumerate() {
            let pattern = &SEARCH_PATTERNS[i];
            let pat_bytes = pattern.as_bytes();
            let mut search_from = 0;

            while let Some(rel_pos) = formula_upper[search_from..].find(pattern.as_str()) {
                let abs_pos = search_from + rel_pos;

                // Word boundary check: reject XLOOKUP→LOOKUP, VLOOKUP→LOOKUP, etc.
                let is_word_start =
                    abs_pos == 0 || !formula_bytes[abs_pos - 1].is_ascii_alphabetic();

                if is_word_start {
                    // Check if inside a string literal
                    if is_inside_string(formula_upper, abs_pos) {
                        search_from = abs_pos + pat_bytes.len();
                        continue;
                    }

                    let should_flag = match exact_match_arg_idx {
                        None => true, // LOOKUP: always approximate
                        Some(arg_idx) => {
                            let paren_start = abs_pos + func_name.len();
                            !has_exact_match_arg(formula_upper, paren_start, arg_idx)
                        }
                    };

                    if should_flag {
                        out.push(Violation::with_data(
                            RuleId::Calc204,
                            ViolationScope::Cell(sheet.sheet_index, CellReference::new(at.0, at.1)),
                            ApproximateLookupData {
                                function: func_name,
                            },
                            Severity::Warning,
                        ));
                        break; // One violation per function per cell
                    }
                }

                search_from = abs_pos + pat_bytes.len();
            }
        }
    }
}

/// Check if the argument at `arg_idx` in a function call starting at
/// `paren_pos` of an uppercased formula indicates exact match (FALSE or 0).
///
/// Returns `true` if the argument exists and is an exact-match indicator.
fn has_exact_match_arg(formula_upper: &str, paren_pos: usize, arg_idx: usize) -> bool {
    let Some(args) = extract_args(formula_upper, paren_pos) else {
        return false; // Malformed formula, flag it
    };

    if arg_idx >= args.len() {
        return false; // Argument missing → defaults to approximate
    }

    // Exact-match indicators
    matches!(args[arg_idx].trim(), "FALSE" | "0")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::LinterContext;
    use std::collections::HashMap;

    fn make_formula_cell(row: u32, col: u32, formula: &str) -> ((u32, u32), Cell) {
        (
            (row, col),
            Cell {
                formula: Some(<Box<str>>::from(formula)),
                num_fmt: None,
                value: CellValue::Empty,
                is_array: false,
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

    fn run_rule(sheet: &Sheet) -> Vec<Violation> {
        let mut rule = ApproximateLookupRule;
        let mut ctx = LinterContext::default();
        let mut violations = Vec::new();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(sheet, at, cell, &mut ctx, &mut violations);
        }
        violations
    }

    // --- VLOOKUP tests ---

    #[test]
    fn test_vlookup_missing_4th_arg() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "VLOOKUP(A1,B:C,2)")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 1);
    }

    #[test]
    fn test_vlookup_with_false() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "VLOOKUP(A1,B:C,2,FALSE)")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 0);
    }

    #[test]
    fn test_vlookup_with_zero() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "VLOOKUP(A1,B:C,2,0)")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 0);
    }

    #[test]
    fn test_vlookup_with_true() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "VLOOKUP(A1,B:C,2,TRUE)")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 1);
    }

    // --- HLOOKUP tests ---

    #[test]
    fn test_hlookup_missing_4th_arg() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "HLOOKUP(A1,1:3,2)")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 1);
    }

    // --- MATCH tests ---

    #[test]
    fn test_match_missing_3rd_arg() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "MATCH(A1,B:B)")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 1);
    }

    #[test]
    fn test_match_with_zero() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "MATCH(A1,B:B,0)")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 0);
    }

    #[test]
    fn test_match_with_one() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "MATCH(A1,B:B,1)")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 1);
    }

    #[test]
    fn test_match_with_negative_one() {
        // match_type=-1 is reverse approximate match — should flag
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "MATCH(A1,B:B,-1)")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 1);
    }

    // --- LOOKUP tests ---

    #[test]
    fn test_lookup_unconditional() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "LOOKUP(A1,B1:B10,C1:C10)")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 1);
    }

    // --- Edge cases ---

    #[test]
    fn test_nested_vlookup() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "IF(A1,VLOOKUP(A1,B:C,2),0)")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 1);
    }

    #[test]
    fn test_xlookup_no_false_positive() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "XLOOKUP(A1,B1:B10,C1:C10)")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 0);
    }

    #[test]
    fn test_string_literal_no_false_positive() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "\"VLOOKUP(\"")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 0);
    }

    #[test]
    fn test_lowercase_vlookup() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "vlookup(A1,B:C,2)")]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 1);
    }

    #[test]
    fn test_vlookup_nested_function_arg() {
        // 4th arg is a nested function call, not FALSE/0
        let sheet = make_sheet(vec![make_formula_cell(
            0,
            0,
            "VLOOKUP(A1,B:C,2,IF(D1,TRUE,FALSE))",
        )]);
        let v = run_rule(&sheet);
        assert_eq!(v.len(), 1, "Nested function as match_type arg should flag");
    }

    #[test]
    fn test_vlookup_second_call_missing_exact_match() {
        // Regression: first VLOOKUP has FALSE (exact), second omits it.
        // The unconditional break bug would skip the second call entirely.
        let sheet = make_sheet(vec![make_formula_cell(
            0,
            0,
            "IF(cond,VLOOKUP(A1,B:C,2,FALSE),VLOOKUP(A1,D:E,2))",
        )]);
        let v = run_rule(&sheet);
        assert_eq!(
            v.len(),
            1,
            "Second VLOOKUP missing exact-match must be flagged"
        );
    }
}

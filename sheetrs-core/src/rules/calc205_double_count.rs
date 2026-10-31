//! CALC205: Double Count
//!
//! When two arguments of one summary function cover the same cells, those cells are counted twice:
//! `=SUM(A1:A10,A5:A15)` includes `A5:A10` in both. The total looks plausible and is too large, and
//! nothing in the spreadsheet signals it.
//!
//! # Reports
//!
//! One violation per formula cell (unit: the cell) in which two arguments of the same call of
//! `SUM`, `COUNT`, `COUNTA` or `AVERAGE` are cell or range references on the same sheet whose rectangles
//! share at least one cell. Single cells count as one-cell ranges (`SUM(A1,A1:A3)`). Only the
//! first overlapping pair of the first offending call is reported, and the message names the
//! function and both arguments. Names are matched case-insensitively, so sheet qualifiers are too.
//!
//! # Does not report
//!
//! - Ranges that merely touch (`A1:A5` and `A6:A10`).
//! - Ranges on different sheets, or with different sheet qualifiers: the qualifier text must be
//!   identical. An unqualified `A1:A5` is therefore not compared with `Sheet1!A1:A5` on the same
//!   sheet.
//! - Arguments that are not a plain reference (nested calls, names, whole columns such as `A:A`,
//!   constants); they are skipped, the remaining ones are still compared.
//! - Functions whose name merely ends in a listed one (`DSUM`) and names inside string constants.
//! - Overlap across separate calls (`=SUM(A1:A10)+SUM(A5:A15)`); each call is checked alone.
//! - The conditional functions (`SUMIF`, `COUNTIF`, `AVERAGEIF` and their `*IFS` forms): their
//!   arguments are ranges, criteria and sum ranges, not values added together, so
//!   `COUNTIF(A1:A10,A5)` (a criterion inside the range) and `SUMIF(A1:A10,">5",A1:A10)` are
//!   correct.
//! - `SUMPRODUCT`: its arguments are multiplied, not added (`SUMPRODUCT(A1:A3,A1:A3)`).
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always `Warning`.

use super::helpers::{extract_args, is_inside_string, parse_formula_range};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::sync::LazyLock;

/// Functions whose every argument contributes values, so an overlap counts cells twice.
const SUM_FAMILY: &[&str] = &["SUM", "COUNT", "COUNTA", "AVERAGE"];

/// Pre-computed search patterns: "FUNC(" for each SUM-family function.
static SEARCH_PATTERNS: LazyLock<Vec<String>> =
    LazyLock::new(|| SUM_FAMILY.iter().map(|name| format!("{name}(")).collect());

/// Rule that identifies double count issues in formulas.
pub struct DoubleCountRule;

/// Incident data for CALC205.
#[derive(Debug)]
pub struct DoubleCountData {
    /// The function containing overlapping ranges.
    pub function: String,
    /// First overlapping range argument (display string).
    pub range_a: String,
    /// Second overlapping range argument (display string).
    pub range_b: String,
}

impl ViolationData for DoubleCountData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!(
            "Overlapping ranges in {}(): {} and {} overlap, causing double-counting.",
            self.function, self.range_a, self.range_b
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Check if two rectangular ranges overlap.
fn ranges_overlap(a: (u32, u32, u32, u32), b: (u32, u32, u32, u32)) -> bool {
    let (a_sr, a_sc, a_er, a_ec) = a;
    let (b_sr, b_sc, b_er, b_ec) = b;
    a_sr <= b_er && b_sr <= a_er && a_sc <= b_ec && b_sc <= a_ec
}

impl WalkerRule for DoubleCountRule {
    fn id(&self) -> RuleId {
        RuleId::Calc205
    }

    fn name(&self) -> &'static str {
        "Double Count"
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

        for (i, pattern) in SEARCH_PATTERNS.iter().enumerate() {
            let pat_bytes = pattern.as_bytes();
            let mut search_from = 0;

            while let Some(rel_pos) = formula_upper[search_from..].find(pattern.as_str()) {
                let abs_pos = search_from + rel_pos;

                // Word boundary check: reject e.g. DSUM matching SUM
                let is_word_start =
                    abs_pos == 0 || !formula_bytes[abs_pos - 1].is_ascii_alphabetic();

                if is_word_start && !is_inside_string(formula_upper, abs_pos) {
                    let paren_pos = abs_pos + SUM_FAMILY[i].len();
                    if let Some(overlap) =
                        find_overlapping_args(formula_upper, paren_pos, SUM_FAMILY[i])
                    {
                        out.push(Violation::with_data(
                            RuleId::Calc205,
                            ViolationScope::Cell(sheet.sheet_index, CellReference::new(at.0, at.1)),
                            overlap,
                            Severity::Warning,
                        ));
                        return; // One violation per cell
                    }
                }

                search_from = abs_pos + pat_bytes.len();
            }
        }
    }
}

/// A parsed range argument with its sheet qualifier and coordinates.
struct ParsedRange {
    display: String,
    sheet_qualifier: String,
    coords: (u32, u32, u32, u32),
}

/// Extract range arguments from a function call and check all pairs for overlap.
///
/// Returns `Some(DoubleCountData)` with the first overlapping pair found,
/// or `None` if no overlaps exist.
fn find_overlapping_args(
    formula: &str,
    paren_pos: usize,
    func_name: &str,
) -> Option<DoubleCountData> {
    let args = extract_args(formula, paren_pos)?;

    // Parse each argument as a range reference; skip non-parseable args.
    // Track sheet qualifier to avoid false positives on cross-sheet refs.
    let parsed_ranges: Vec<ParsedRange> = args
        .iter()
        .filter_map(|arg| {
            let trimmed = arg.trim();
            let (sheet_qualifier, range_part) = trimmed
                .rsplit_once('!')
                .map_or(("", trimmed), |(sheet, range)| (sheet, range));
            parse_formula_range(range_part).map(|coords| ParsedRange {
                display: trimmed.to_string(),
                sheet_qualifier: sheet_qualifier.to_string(),
                coords,
            })
        })
        .collect();

    // Check all pairs — only compare ranges on the same sheet
    for i in 0..parsed_ranges.len() {
        for j in (i + 1)..parsed_ranges.len() {
            if parsed_ranges[i].sheet_qualifier != parsed_ranges[j].sheet_qualifier {
                continue; // Different sheets cannot overlap
            }
            if ranges_overlap(parsed_ranges[i].coords, parsed_ranges[j].coords) {
                return Some(DoubleCountData {
                    function: func_name.to_string(),
                    range_a: parsed_ranges[i].display.clone(),
                    range_b: parsed_ranges[j].display.clone(),
                });
            }
        }
    }

    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;

    fn make_formula_cell(row: u32, col: u32, formula: &str) -> ((u32, u32), Cell) {
        (
            (row, col),
            Cell {
                formula: Some(<Box<str>>::from(formula)),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
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
        let mut rule = DoubleCountRule;
        let mut ctx = LinterContext::default();
        let mut violations = Vec::new();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(sheet, at, cell, &mut ctx, &mut violations);
        }
        violations
    }

    #[test]
    fn test_overlapping_sum_ranges() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "SUM(A1:A10,A5:A15)")]);
        let violations = run_rule(&sheet);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Calc205);
        let data = violations[0].data::<DoubleCountData>().unwrap();
        assert_eq!(data.function, "SUM");
    }

    #[test]
    fn test_non_overlapping_sum_ranges() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "SUM(A1:A10,B1:B10)")]);
        let violations = run_rule(&sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_adjacent_non_overlapping() {
        // A1:A5 and A6:A10 are adjacent but do NOT overlap
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "SUM(A1:A5,A6:A10)")]);
        let violations = run_rule(&sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_single_argument() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "SUM(A1:A10)")]);
        let violations = run_rule(&sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_count_overlapping() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "COUNT(A1:A10,A5:A15)")]);
        let violations = run_rule(&sheet);

        assert_eq!(violations.len(), 1);
        let data = violations[0].data::<DoubleCountData>().unwrap();
        assert_eq!(data.function, "COUNT");
    }

    #[test]
    fn test_nested_function_no_false_positive() {
        // IF(cond,A1:A10) is not parseable as a range → skip, no false positive
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "SUM(IF(cond,A1:A10),A1:A10)")]);
        let violations = run_rule(&sheet);

        // First arg is "IF(cond,A1:A10)" which is not a valid range → only 1 range parsed → no overlap
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
                value: CellValue::Number(42.0),
            },
        );
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cell_map.into(),
            ..Default::default()
        };
        let violations = run_rule(&sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_case_insensitive() {
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "sum(A1:A10,A5:A15)")]);
        let violations = run_rule(&sheet);

        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_case_insensitive_non_ascii_sheet_qualifier() {
        let sheet = make_sheet(vec![make_formula_cell(
            0,
            0,
            "SUM('Año'!A1:A10,'AÑO'!A5:A15)",
        )]);
        let violations = run_rule(&sheet);

        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn conditional_functions_and_sumproduct_are_not_double_counts() {
        for f in [
            "COUNTIF(A1:A10,A5)",
            "SUMIF(A1:A10,\">5\",A1:A10)",
            "AVERAGEIFS(A1:A10,B1:B10,B5,A1:A10)",
            "SUMPRODUCT(A1:A3,A1:A3)",
        ] {
            let sheet = make_sheet(vec![make_formula_cell(0, 5, f)]);
            assert!(run_rule(&sheet).is_empty(), "{f}");
        }
    }

    #[test]
    fn test_dsum_no_false_positive() {
        // DSUM should not match SUM
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "DSUM(A1:A10,A5:A15)")]);
        let violations = run_rule(&sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_cross_sheet_no_false_positive() {
        // Same coordinates on different sheets should NOT overlap
        let sheet = make_sheet(vec![make_formula_cell(
            0,
            0,
            "SUM(Sheet1!A1:A10,Sheet2!A1:A10)",
        )]);
        let violations = run_rule(&sheet);

        assert_eq!(violations.len(), 0);
    }
}

//! CALC203: Double Operator detection
//!
//! Description: Flags redundant operator sequences (e.g., "++", "--") indicating potential typos.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use regex::Regex;
use std::sync::LazyLock;

/// Regex matching truly redundant consecutive arithmetic operators.
/// Matches pairs like `++`, `--`, `+-`, `-+`, `**`, `//`, `^^`, `*/`, `/*`, etc.
/// Excludes valid unary +/- after binary operators `*`, `/`, `^` (e.g. `A1*-B1`).
static DOUBLE_OP_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"[+\-]\s*[+\-*/^]|[*/^]\s*[*/^]").unwrap());

/// Regex for stripping string literals from formulas to avoid false positives.
static STRING_LITERAL_RE: LazyLock<Regex> = LazyLock::new(|| Regex::new(r#""[^"]*""#).unwrap());

/// Rule that identifies double operators in formulas.
pub struct DoubleOperatorRule {
    /// Whether to allow `--` as intentional double-negative coercion.
    allow_double_negative: bool,
}

impl DoubleOperatorRule {
    /// Create a new instance with configuration.
    pub fn new(config: &LinterConfig) -> Self {
        Self {
            allow_double_negative: config
                .get_param_bool("calc203_allow_double_negative", None)
                .unwrap_or(false),
        }
    }
}

/// Incident data for CALC203.
#[derive(Debug)]
pub struct DoubleOperatorData {
    /// The matched operator pair (e.g., "++", "--").
    pub matched_ops: String,
}

impl ViolationData for DoubleOperatorData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!(
            "Double operator \"{}\" found in formula. This may indicate a typo.",
            self.matched_ops
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for DoubleOperatorRule {
    fn id(&self) -> RuleId {
        RuleId::Calc203
    }

    fn name(&self) -> &'static str {
        "Double Operator"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Calculations
    }

    fn on_cell(&mut self, sheet: &Sheet, cell: &Cell, _ctx: &mut LinterContext) -> Vec<Violation> {
        let Some(formula) = cell.as_formula() else {
            return Vec::new();
        };

        // Strip string literals to avoid false positives in text constants
        let cleaned = STRING_LITERAL_RE.replace_all(formula, "");

        let mut violations = Vec::new();

        for m in DOUBLE_OP_RE.find_iter(&cleaned) {
            let matched = m.as_str();
            // Extract just the operator characters (strip whitespace)
            let ops: String = matched.chars().filter(|c| !c.is_whitespace()).collect();

            // Skip `--` if allow_double_negative is enabled
            if self.allow_double_negative && ops == "--" {
                continue;
            }

            violations.push(Violation::with_data(
                RuleId::Calc203,
                ViolationScope::Cell(sheet.sheet_index, CellReference::new(cell.row, cell.col)),
                DoubleOperatorData { matched_ops: ops },
                Severity::Warning,
            ));
            // One violation per cell is sufficient
            break;
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

    fn run_rule(rule: &mut DoubleOperatorRule, sheet: &Sheet) -> Vec<Violation> {
        let mut ctx = LinterContext::default();
        let mut violations = Vec::new();
        for cell in sheet.all_cells() {
            violations.extend(rule.on_cell(sheet, cell, &mut ctx));
        }
        violations
    }

    #[test]
    fn test_double_plus() {
        let config = LinterConfig::default();
        let mut rule = DoubleOperatorRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "A1++B1")]);
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Calc203);
        let data = violations[0].data::<DoubleOperatorData>().unwrap();
        assert_eq!(data.matched_ops, "++");
    }

    #[test]
    fn test_double_minus() {
        let config = LinterConfig::default();
        let mut rule = DoubleOperatorRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "A1--B1")]);
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_double_minus_allowed() {
        let mut config = LinterConfig::default();
        config.global.params.insert(
            "calc203_allow_double_negative".to_string(),
            toml::Value::Boolean(true),
        );
        let mut rule = DoubleOperatorRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "SUMPRODUCT(--(A1:A10>5))")]);
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_operators_with_whitespace() {
        let config = LinterConfig::default();
        let mut rule = DoubleOperatorRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "A1 + + B1")]);
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_string_literal_not_flagged() {
        let config = LinterConfig::default();
        let mut rule = DoubleOperatorRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "CONCAT(\"++\", A1)")]);
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_single_operator_ok() {
        let config = LinterConfig::default();
        let mut rule = DoubleOperatorRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "A1+B1")]);
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_star_slash() {
        let config = LinterConfig::default();
        let mut rule = DoubleOperatorRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "A1*/B1")]);
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 1);
        let data = violations[0].data::<DoubleOperatorData>().unwrap();
        assert_eq!(data.matched_ops, "*/");
    }

    #[test]
    fn test_unary_minus_ok() {
        let config = LinterConfig::default();
        let mut rule = DoubleOperatorRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "-A1")]);
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_multiply_by_negative_ok() {
        let config = LinterConfig::default();
        let mut rule = DoubleOperatorRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "A1*-B1")]);
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_divide_by_negative_ok() {
        let config = LinterConfig::default();
        let mut rule = DoubleOperatorRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "A1/-1")]);
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_power_of_negative_ok() {
        let config = LinterConfig::default();
        let mut rule = DoubleOperatorRule::new(&config);
        let sheet = make_sheet(vec![make_formula_cell(0, 0, "A1^-2")]);
        let violations = run_rule(&mut rule, &sheet);

        assert_eq!(violations.len(), 0);
    }
}

//! CPX504: Deep IF nesting (primary emitter of the CPX5xx formula-complexity chain)
//!
//! Deeply nested `IF` calls are hard to read, test and change; a lookup table, `IFS` or `SWITCH`
//! usually states the same logic better. This module also evaluates, for each formula cell, the whole
//! chain CPX504 to CPX509, and is the only place that emits those violations. The modules
//! `cpx505_deep_formula_nesting`, `cpx506_many_operations`, `cpx508_many_references` and
//! `cpx509_long_formula` are no-op walker rules that exist only so that each rule can be enabled or
//! disabled; each one's results are the same whichever of them are enabled.
//! [`cpx507_multiple_sheet_ref`](super::cpx507_multiple_sheet_ref) is independent of the chain.
//!
//! # Reports
//! One violation per formula cell (scope: cell) whose IF nesting depth is greater than
//! `max_if_nesting`. The cell is then not examined further. A cell that does not exceed it is passed
//! to the next check of the chain, in this order, and the first that fires ends the examination:
//!
//! 1. CPX504: IF depth greater than `max_if_nesting`.
//! 2. CPX505: parenthesis depth greater than `max_formula_nesting`.
//! 3. CPX506: operators greater than `max_operations`.
//! 4. CPX508: references greater than `max_references`, only when CPX506 did not fire.
//! 5. CPX509: length greater than `max_formula_length`, only when CPX505 did not fire.
//!
//! CPX508 and CPX509 do not exclude each other, so one cell can carry both; every other outcome is
//! a single violation. A cell that fires an earlier check does not also report a later one: the
//! earlier finding is the more specific advice, and one cell is not buried under overlapping
//! complaints.
//!
//! The IF depth is the largest number of `IF(` calls open at the same time in the upper-cased
//! formula text. Only a call of `IF` counts (not the end of `SUMIF(`, `COUNTIF(` or a prefixed name),
//! a `)` closes the call it belongs to (the parenthesis that ends `SUM(...)` inside an `IF` does not
//! end the `IF`), and string literals and quoted sheet names are not read as code.
//!
//! # Does not report
//! Non-formula cells, and IF depth at or below the limit. Nesting of other functions is not
//! measured here; CPX505 does that.
//!
//! # Parameters
//! - `max_if_nesting` (integer, default `5`); the parameters of the rest of the chain
//!   (`max_formula_nesting` `5`, `max_operations` `8`, `max_references` `10`, `max_formula_length`
//!   `255`) are read here too. All are global.
//!
//! # Severity
//! Always `Warning`, for every rule of the chain.

use super::helpers::mask_quoted;
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};

/// Rule that detects deeply nested IF statements and serves as the
/// primary emitter for the CPX5xx formula-complexity dependency chain.
pub struct DeepIfNestingRule {
    /// Maximum allowed IF nesting depth (CPX504).
    max_if_nesting: usize,
    /// Maximum allowed parenthesis nesting depth (CPX505).
    max_nesting: usize,
    /// Maximum allowed operator count (CPX506).
    max_operations: usize,
    /// Maximum allowed cell references (CPX508).
    max_references: usize,
    /// Maximum allowed formula length in characters (CPX509).
    max_formula_length: usize,
}

impl DeepIfNestingRule {
    /// Create a new instance from configuration.
    pub fn new(config: &LinterConfig) -> Self {
        Self {
            max_if_nesting: config.get_param_int("max_if_nesting", None).unwrap_or(5) as usize,
            max_nesting: config
                .get_param_int("max_formula_nesting", None)
                .unwrap_or(5) as usize,
            max_operations: config.get_param_int("max_operations", None).unwrap_or(8) as usize,
            max_references: config.get_param_int("max_references", None).unwrap_or(10) as usize,
            max_formula_length: config
                .get_param_int("max_formula_length", None)
                .unwrap_or(255) as usize,
        }
    }
}

impl Default for DeepIfNestingRule {
    fn default() -> Self {
        Self {
            max_if_nesting: 5,
            max_nesting: 5,
            max_operations: 8,
            max_references: 10,
            max_formula_length: 255,
        }
    }
}

// ---------------------------------------------------------------------------
// Compact ViolationData structs — incident-only, no setup values
// ---------------------------------------------------------------------------

/// Incident data for CPX504 (deep IF nesting).
#[derive(Debug)]
pub struct DeepIfNestingData {
    /// Detected IF nesting depth.
    pub depth: usize,
}

impl ViolationData for DeepIfNestingData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!(
            "Deeply nested IF statements ({} levels). Consider lookup tables or IFS.",
            self.depth
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Incident data for CPX505 (deep formula nesting).
#[derive(Debug)]
pub struct DeepFormulaNestingData {
    /// Detected parenthesis nesting depth.
    pub depth: usize,
}

impl ViolationData for DeepFormulaNestingData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!(
            "Formula with deep nesting ({} levels). Consider simplifying.",
            self.depth
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Incident data for CPX506 (many operations).
#[derive(Debug)]
pub struct ManyOperationsData {
    /// Detected operator count.
    pub count: usize,
}

impl ViolationData for ManyOperationsData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!("Formula with many operations ({}).", self.count)
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Incident data for CPX508 (many references).
#[derive(Debug)]
pub struct ManyReferencesData {
    /// Detected reference count.
    pub count: usize,
}

impl ViolationData for ManyReferencesData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!("Formula with many references ({}).", self.count)
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Incident data for CPX509 (long formula).
#[derive(Debug)]
pub struct LongFormulaData {
    /// Detected formula length in characters.
    pub length: usize,
}

impl ViolationData for LongFormulaData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!("Long formula ({} characters).", self.length)
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

// ---------------------------------------------------------------------------
// Analysis helpers
// ---------------------------------------------------------------------------

/// Count the maximum nesting depth of IF statements in an uppercased formula.
fn count_if_nesting(formula_upper: &str) -> usize {
    let masked = mask_quoted(formula_upper);
    let bytes = masked.as_bytes();
    // One entry per open parenthesis: whether it opens the arguments of an `IF`.
    let mut open: Vec<bool> = Vec::new();
    let mut depth = 0usize;
    let mut max_depth = 0usize;
    for (i, &b) in bytes.iter().enumerate() {
        match b {
            b'(' => {
                let is_if = bytes[..i].ends_with(b"IF")
                    && !bytes[..i - 2]
                        .last()
                        .is_some_and(|p| p.is_ascii_alphanumeric() || *p == b'_' || *p == b'.');
                open.push(is_if);
                if is_if {
                    depth += 1;
                    max_depth = max_depth.max(depth);
                }
            }
            b')' if open.pop() == Some(true) => depth -= 1,
            _ => {}
        }
    }
    max_depth
}

/// Calculate the maximum parenthesis nesting depth of a formula.
fn calculate_nesting_depth(formula: &str) -> usize {
    let mut max_depth: usize = 0;
    let mut current_depth: usize = 0;

    for ch in mask_quoted(formula).chars() {
        match ch {
            '(' => {
                current_depth += 1;
                max_depth = max_depth.max(current_depth);
            }
            ')' => {
                current_depth = current_depth.saturating_sub(1);
            }
            _ => {}
        }
    }

    max_depth
}

/// Count arithmetic and logical operators, skipping string literals.
///
/// Recognised: `+`, `-`, `*`, `/`, `^`, `&`, `=`, `<`, `>`, `<=`, `>=`, `<>`.
/// The leading `=` of a formula is excluded.
fn count_operators(formula: &str) -> usize {
    let mut count: usize = 0;
    let bytes = formula.as_bytes();
    let len = bytes.len();
    let mut i = if !bytes.is_empty() && bytes[0] == b'=' {
        1
    } else {
        0
    };
    let mut in_string = false;

    while i < len {
        let ch = bytes[i];

        if ch == b'"' {
            in_string = !in_string;
            i += 1;
            continue;
        }
        if in_string {
            i += 1;
            continue;
        }

        match ch {
            b'+' | b'-' | b'*' | b'/' | b'^' | b'&' => count += 1,
            b'<' | b'>' => {
                count += 1;
                if i + 1 < len && matches!(bytes[i + 1], b'=' | b'>') {
                    i += 1;
                }
            }
            b'=' => count += 1,
            _ => {}
        }
        i += 1;
    }

    count
}

/// Count cell/range references (e.g. `A1`, `$B$2`, `AB12`).
///
/// Skips string literals and function names.
fn count_references(formula: &str) -> usize {
    let mut count: usize = 0;
    let bytes = formula.as_bytes();
    let len = bytes.len();
    let mut i = if !bytes.is_empty() && bytes[0] == b'=' {
        1
    } else {
        0
    };
    let mut in_string = false;

    while i < len {
        if bytes[i] == b'"' {
            in_string = !in_string;
            i += 1;
            continue;
        }
        if in_string {
            i += 1;
            continue;
        }

        let start = i;

        if i < len && bytes[i] == b'$' {
            i += 1;
        }
        let col_start = i;
        while i < len && bytes[i].is_ascii_alphabetic() {
            i += 1;
        }
        let col_len = i - col_start;
        if col_len == 0 || col_len > 3 {
            i = start + 1;
            continue;
        }

        if i < len && bytes[i] == b'$' {
            i += 1;
        }
        let row_start = i;
        while i < len && bytes[i].is_ascii_digit() {
            i += 1;
        }
        if i == row_start {
            i = start + 1;
            continue;
        }
        if i < len && (bytes[i].is_ascii_alphanumeric() || bytes[i] == b'_') {
            i = start + 1;
            continue;
        }

        count += 1;
    }

    count
}

// ---------------------------------------------------------------------------
// WalkerRule implementation
// ---------------------------------------------------------------------------

impl WalkerRule for DeepIfNestingRule {
    fn id(&self) -> RuleId {
        RuleId::Cpx504
    }

    fn name(&self) -> &'static str {
        "Deep IF Nesting"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Complexity
    }

    /// Evaluates the CPX5xx dependency chain for one formula cell.
    fn on_cell(
        &mut self,
        sheet: &Sheet,
        at: (u32, u32),
        cell: &Cell,
        ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        let Some(formula) = cell.as_formula() else {
            return;
        };

        let scope = ViolationScope::Cell(sheet.sheet_index, CellReference::new(at.0, at.1));

        // --- Dependency cascade CPX504 → CPX505 → CPX506 → CPX508 / CPX509 ---

        let if_depth = count_if_nesting(
            ctx.formula_upper(sheet, at, cell)
                .expect("formula cell has an uppercased formula"),
        );
        if if_depth > self.max_if_nesting {
            out.push(Violation::with_data(
                RuleId::Cpx504,
                scope,
                DeepIfNestingData { depth: if_depth },
                Severity::Warning,
            ));
            return;
        }

        let nesting = calculate_nesting_depth(formula);
        if nesting > self.max_nesting {
            out.push(Violation::with_data(
                RuleId::Cpx505,
                scope,
                DeepFormulaNestingData { depth: nesting },
                Severity::Warning,
            ));
            return;
        }

        let ops = count_operators(formula);
        if ops > self.max_operations {
            out.push(Violation::with_data(
                RuleId::Cpx506,
                scope,
                ManyOperationsData { count: ops },
                Severity::Warning,
            ));
            return;
        }

        // CPX508 — only if CPX506 did not fire
        let refs = count_references(formula);
        if refs > self.max_references {
            out.push(Violation::with_data(
                RuleId::Cpx508,
                scope.clone(),
                ManyReferencesData { count: refs },
                Severity::Warning,
            ));
        }

        // CPX509 — only if CPX505 did not fire
        // The visible length is never above the raw one: the formula is parsed only when the raw
        // length passes the maximum.
        let length = formula.chars().count();
        let length = if length > self.max_formula_length {
            sheetrs_formula::calc_print::visible_length(formula)
        } else {
            length
        };
        if length > self.max_formula_length {
            out.push(Violation::with_data(
                RuleId::Cpx509,
                scope,
                LongFormulaData { length },
                Severity::Warning,
            ));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::collect_violations;
    use std::collections::HashMap;

    fn make_sheet(formula: &str) -> Sheet {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from(formula)),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );
        Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            used_range: Some((1, 1)),
            ..Default::default()
        }
    }

    fn run_violations(sheet: &Sheet, rule: &mut DeepIfNestingRule) -> Vec<Violation> {
        let mut ctx = LinterContext::default();
        sheet
            .cells
            .iter()
            .flat_map(|(at, c)| collect_violations(|out| rule.on_cell(sheet, at, c, &mut ctx, out)))
            .collect()
    }

    #[test]
    fn test_violations_have_cell_scope() {
        let long = "=".to_string() + &"A1+".repeat(100);
        let mut sheet = make_sheet(&long);
        sheet.sheet_index = 2;
        let mut rule = DeepIfNestingRule {
            max_operations: 999,
            ..Default::default()
        };
        let violations = run_violations(&sheet, &mut rule);
        let ids: Vec<_> = violations.iter().map(|v| v.rule_id).collect();
        assert_eq!(ids, [RuleId::Cpx508, RuleId::Cpx509]);
        for v in &violations {
            assert_eq!(v.scope, ViolationScope::Cell(2, CellReference::new(0, 0)));
        }
    }

    #[test]
    fn test_cpx504_suppresses_others() {
        let sheet = make_sheet("=IF(A1,IF(B1,IF(C1,IF(D1,IF(E1,IF(F1,1,0),0),0),0),0),0)");
        let mut rule = DeepIfNestingRule::default();
        let violations = run_violations(&sheet, &mut rule);
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Cpx504);
        let data = violations[0].data::<DeepIfNestingData>().unwrap();
        assert_eq!(data.depth, 6);
    }

    #[test]
    fn test_cpx505_suppresses_downstream() {
        let sheet = make_sheet("=SUM(ABS(MIN(MAX(ROUND(AVERAGE(A1),2),0),100),0),0)");
        let mut rule = DeepIfNestingRule::default();
        let violations = run_violations(&sheet, &mut rule);
        assert!(violations.iter().any(|v| v.rule_id == RuleId::Cpx505));
        assert!(!violations.iter().any(|v| v.rule_id == RuleId::Cpx506));
    }

    #[test]
    fn test_cpx506_suppresses_cpx508() {
        let sheet = make_sheet("=A1+B1*C1-D1/E1^F1&G1+H1+I1+J1");
        let mut rule = DeepIfNestingRule::default();
        let violations = run_violations(&sheet, &mut rule);
        assert!(violations.iter().any(|v| v.rule_id == RuleId::Cpx506));
        assert!(!violations.iter().any(|v| v.rule_id == RuleId::Cpx508));
    }

    #[test]
    fn test_cpx508_fires_when_cpx506_does_not() {
        let sheet = make_sheet("=A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1");
        let mut rule = DeepIfNestingRule {
            max_operations: 15,
            ..Default::default()
        };
        let violations = run_violations(&sheet, &mut rule);
        assert!(violations.iter().any(|v| v.rule_id == RuleId::Cpx508));
    }

    #[test]
    fn test_cpx509_fires_alone() {
        let long = "=".to_string() + &"A1+".repeat(100);
        let sheet = make_sheet(&long);
        let mut rule = DeepIfNestingRule {
            max_operations: 999,
            max_references: 999,
            ..Default::default()
        };
        let violations = run_violations(&sheet, &mut rule);
        assert!(violations.iter().any(|v| v.rule_id == RuleId::Cpx509));
    }

    #[test]
    fn test_cpx509_counts_characters() {
        let within = format!("=\"{}\"", "ñ".repeat(200));
        assert!(run_violations(&make_sheet(&within), &mut DeepIfNestingRule::default()).is_empty());

        let over = format!("=\"{}\"", "ñ".repeat(254));
        let violations = run_violations(&make_sheet(&over), &mut DeepIfNestingRule::default());
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].data::<LongFormulaData>().unwrap().length, 257);
    }

    #[test]
    fn test_violation_data_downcast() {
        let sheet = make_sheet("=A1+B1*C1-D1/E1^F1&G1+H1+I1+J1");
        let mut rule = DeepIfNestingRule::default();
        let violations = run_violations(&sheet, &mut rule);
        let v = violations
            .iter()
            .find(|v| v.rule_id == RuleId::Cpx506)
            .unwrap();
        let data = v.data::<ManyOperationsData>().unwrap();
        assert!(data.count > 8);
    }

    #[test]
    fn test_if_counting() {
        assert_eq!(count_if_nesting("=A1+B1"), 0);
        assert_eq!(count_if_nesting("=IF(A1,1,0)"), 1);
        assert_eq!(count_if_nesting("=IF(A1,IF(B1,1,0),0)"), 2);
    }

    #[test]
    fn test_operator_counting() {
        assert_eq!(count_operators("=A1+B1"), 1);
        assert_eq!(count_operators("=IF(A1<=B1,1,0)"), 1);
        assert_eq!(count_operators(r#"=A1&"a+b""#), 1);
    }

    #[test]
    fn test_reference_counting() {
        assert_eq!(count_references("=A1+B1"), 2);
        assert_eq!(count_references("=SUM(1,2,3)"), 0);
    }

    #[test]
    fn test_no_violations_simple_formula() {
        let sheet = make_sheet("=A1+B1");
        let mut rule = DeepIfNestingRule::default();
        let violations = run_violations(&sheet, &mut rule);
        assert!(violations.is_empty());
    }

    #[test]
    fn test_count_if_nesting_non_ascii() {
        // Non-ASCII chars like Ó (2 bytes in UTF-8) must not cause a panic.
        assert_eq!(count_if_nesting(r#""INFORMACIÓN"&IF(A1,1,0)"#), 1);
        assert_eq!(count_if_nesting(r#"IF("AÑO"="",IF(A1,1,0),0)"#), 2);
    }

    #[test]
    fn if_depth_counts_calls_of_if_and_the_parentheses_that_belong_to_them() {
        // SUMIF is not IF.
        assert_eq!(count_if_nesting("=SUMIF(A1:A5,\">1\",B1:B5)"), 0);
        assert_eq!(count_if_nesting("=IF(A1,SUMIF(A1:A5,1,B1:B5),0)"), 1);
        // The parenthesis that closes another function does not close an IF.
        assert_eq!(
            count_if_nesting("=IF(A1,SUM(B1,C1),IF(B1,IF(C1,1,0),0))"),
            3
        );
        // Text is not code.
        assert_eq!(count_if_nesting("=\"IF(IF(IF(\"&IF(A1,1,0)"), 1);
        assert_eq!(count_if_nesting("=IF(A1,\")))\",IF(B1,1,0))"), 2);
        // A prefixed or longer name is not IF.
        assert_eq!(count_if_nesting("=_XLFN.IFS(A1,1,B1,2)+MYIF(A1)"), 0);
    }

    #[test]
    fn paren_depth_ignores_parentheses_in_text() {
        assert_eq!(calculate_nesting_depth("=A1&\"((((\""), 0);
        assert_eq!(calculate_nesting_depth("=SUM((A1+B1)*2)&\")\""), 2);
    }
}

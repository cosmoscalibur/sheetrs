//! CPX504: Primary emitter for the CPX5xx formula-complexity cascade
//!
//! Description: Detects deeply nested IF statements which are hard to maintain and debug.
//!
//! This module is the **primary emitter** for the CPX5xx formula-complexity chain.
//! On each cell it evaluates the full dependency cascade and emits at most one
//! violation per cell following the priority order:
//!
//! 1. **CPX504** — deep IF nesting (>`max_if_nesting`)
//! 2. **CPX505** — deep parenthesis nesting (>`max_formula_nesting`)
//! 3. **CPX506** — many operators (>`max_operations`)
//! 4. **CPX508** — many cell references (>`max_references`)  *(suppressed by CPX506)*
//! 5. **CPX509** — long formula (>`max_formula_length`) *(suppressed by CPX505)*
//!
//! CPX507 (multiple sheet refs) is independent — handled by its own module.
//!
//! The subordinate modules (`cpx505`, `cpx506`, `cpx508`, `cpx509`) are no-op walker
//! shells registered only for configuration enable/disable.

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
    let mut max_depth: usize = 0;
    let mut current_depth: usize = 0;
    let bytes = formula_upper.as_bytes();
    let len = bytes.len();
    let mut i = 0;

    while i < len {
        if i + 3 <= len && bytes[i..i + 3] == *b"IF(" {
            current_depth += 1;
            max_depth = max_depth.max(current_depth);
            i += 3;
        } else if bytes[i] == b')' {
            current_depth = current_depth.saturating_sub(1);
            i += 1;
        } else {
            i += 1;
        }
    }

    max_depth
}

/// Calculate the maximum parenthesis nesting depth of a formula.
fn calculate_nesting_depth(formula: &str) -> usize {
    let mut max_depth: usize = 0;
    let mut current_depth: usize = 0;

    for ch in formula.chars() {
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
    fn on_cell(&mut self, sheet: &Sheet, cell: &Cell, ctx: &mut LinterContext) -> Vec<Violation> {
        let Some(formula) = cell.as_formula() else {
            return Vec::new();
        };

        let scope = ViolationScope::Cell(sheet.sheet_index, CellReference::new(cell.row, cell.col));

        // --- Dependency cascade CPX504 → CPX505 → CPX506 → CPX508 / CPX509 ---

        let if_depth = count_if_nesting(
            ctx.formula_upper(sheet, cell)
                .expect("formula cell has an uppercased formula"),
        );
        if if_depth > self.max_if_nesting {
            return vec![Violation::with_data(
                RuleId::Cpx504,
                scope,
                DeepIfNestingData { depth: if_depth },
                Severity::Warning,
            )];
        }

        let nesting = calculate_nesting_depth(formula);
        if nesting > self.max_nesting {
            return vec![Violation::with_data(
                RuleId::Cpx505,
                scope,
                DeepFormulaNestingData { depth: nesting },
                Severity::Warning,
            )];
        }

        let ops = count_operators(formula);
        if ops > self.max_operations {
            return vec![Violation::with_data(
                RuleId::Cpx506,
                scope,
                ManyOperationsData { count: ops },
                Severity::Warning,
            )];
        }

        let mut violations = Vec::new();

        // CPX508 — only if CPX506 did not fire
        let refs = count_references(formula);
        if refs > self.max_references {
            violations.push(Violation::with_data(
                RuleId::Cpx508,
                scope.clone(),
                ManyReferencesData { count: refs },
                Severity::Warning,
            ));
        }

        // CPX509 — only if CPX505 did not fire
        let length = formula.chars().count();
        if length > self.max_formula_length {
            violations.push(Violation::with_data(
                RuleId::Cpx509,
                scope,
                LongFormulaData { length },
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
    use std::collections::HashMap;

    fn make_sheet(formula: &str) -> Sheet {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from(formula)),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );
        Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 1)),
            ..Default::default()
        }
    }

    fn run_violations(sheet: &Sheet, rule: &mut DeepIfNestingRule) -> Vec<Violation> {
        let mut ctx = LinterContext::default();
        sheet
            .all_cells()
            .flat_map(|c| rule.on_cell(sheet, c, &mut ctx))
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
}

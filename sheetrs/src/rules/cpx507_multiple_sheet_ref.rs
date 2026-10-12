//! CPX507: Multiple Sheet Ref detection
//!
//! Description: Flags formulas referencing data from many different sheets, increasing complexity.
//!
//! This rule is **independent** from the CPX5xx dependency cascade and operates
//! entirely in its own walker.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::HashSet;

/// Rule that identifies formulas referencing many different sheets.
pub struct MultipleSheetRefRule {
    /// Maximum allowed distinct sheet references before flagging.
    max_sheet_refs: usize,
    /// Violations of the current sheet with their `(row, col)`, emitted in
    /// cell order at sheet end because the walker visits cells in hash order.
    pending: Vec<((u32, u32), Violation)>,
}

impl MultipleSheetRefRule {
    /// Create a new instance from configuration.
    pub fn new(config: &LinterConfig) -> Self {
        let max_sheet_refs = config.get_param_int("max_sheet_refs", None).unwrap_or(3) as usize;
        Self {
            max_sheet_refs,
            pending: Vec::new(),
        }
    }
}

impl Default for MultipleSheetRefRule {
    fn default() -> Self {
        Self {
            max_sheet_refs: 3,
            pending: Vec::new(),
        }
    }
}

/// Incident data for CPX507.
#[derive(Debug)]
pub struct MultipleSheetRefData {
    /// Number of distinct sheets referenced.
    pub count: usize,
    /// Cell where the formula resides.
    pub cell: CellReference,
}

impl ViolationData for MultipleSheetRefData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!(
            "Formula references {} distinct sheets at {}.",
            self.count, self.cell
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Extract distinct sheet names referenced in a formula.
///
/// Recognises two patterns:
/// - Unquoted: `SheetName!` where the name is `[A-Za-z0-9_]+`
/// - Quoted: `'Sheet Name'!` (single-quoted, may contain spaces/special chars)
///
/// Content inside double-quoted strings is skipped.
fn count_sheet_refs(formula: &str) -> usize {
    let mut refs: HashSet<&str> = HashSet::new();
    let bytes = formula.as_bytes();
    let len = bytes.len();
    let mut i = 0;
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

        if bytes[i] == b'\'' {
            if let Some(close) = formula[i + 1..].find('\'') {
                let name_end = i + 1 + close;
                if name_end + 1 < len && bytes[name_end + 1] == b'!' {
                    let name = &formula[i + 1..name_end];
                    if !name.is_empty() {
                        refs.insert(name);
                    }
                    i = name_end + 2;
                    continue;
                }
            }
            i += 1;
            continue;
        }

        if bytes[i].is_ascii_alphanumeric() || bytes[i] == b'_' {
            let start = i;
            while i < len && (bytes[i].is_ascii_alphanumeric() || bytes[i] == b'_') {
                i += 1;
            }
            if i < len && bytes[i] == b'!' {
                refs.insert(&formula[start..i]);
                i += 1;
                continue;
            }
            continue;
        }

        i += 1;
    }

    refs.len()
}

impl WalkerRule for MultipleSheetRefRule {
    fn id(&self) -> RuleId {
        RuleId::Cpx507
    }

    fn name(&self) -> &'static str {
        "Multiple Sheet Ref"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Complexity
    }

    fn on_cell(&mut self, sheet: &Sheet, cell: &Cell, _ctx: &mut LinterContext) -> Vec<Violation> {
        if let Some(formula) = cell.as_formula() {
            let count = count_sheet_refs(formula);
            if count >= self.max_sheet_refs {
                self.pending.push((
                    (cell.row, cell.col),
                    Violation::with_data(
                        RuleId::Cpx507,
                        ViolationScope::Sheet(sheet.sheet_index),
                        MultipleSheetRefData {
                            count,
                            cell: CellReference::new(cell.row, cell.col),
                        },
                        Severity::Warning,
                    ),
                ));
            }
        }
        Vec::new()
    }

    fn on_sheet_end(&mut self, _sheet: &Sheet, _ctx: &mut LinterContext) -> Vec<Violation> {
        self.pending.sort_unstable_by_key(|(cell, _)| *cell);
        self.pending.drain(..).map(|(_, v)| v).collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;

    fn run_violations(sheet: &Sheet) -> Vec<Violation> {
        let mut rule = MultipleSheetRefRule::default();
        let mut ctx = LinterContext::default();
        let mut violations: Vec<_> = sheet
            .all_cells()
            .flat_map(|c| rule.on_cell(sheet, c, &mut ctx))
            .collect();
        violations.extend(rule.on_sheet_end(sheet, &mut ctx));
        violations
    }

    #[test]
    fn test_extract_unquoted_refs() {
        assert_eq!(count_sheet_refs("=Sheet1!A1+Sheet2!B2+Sheet3!C3"), 3);
    }

    #[test]
    fn test_extract_quoted_refs() {
        assert_eq!(count_sheet_refs("='My Sheet'!A1+'Other Sheet'!B2"), 2);
    }

    #[test]
    fn test_extract_mixed_refs() {
        assert_eq!(count_sheet_refs("=Sheet1!A1+'My Sheet'!B2+Sheet3!C3"), 3);
    }

    #[test]
    fn test_no_refs() {
        assert_eq!(count_sheet_refs("=SUM(A1:A10)"), 0);
    }

    #[test]
    fn test_refs_inside_string_skipped() {
        assert_eq!(count_sheet_refs(r#"=A1&"Sheet1!A1""#), 0);
    }

    #[test]
    fn test_triggers() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=Sheet1!A1+Sheet2!B2+Sheet3!C3")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let violations = run_violations(&sheet);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Cpx507);
    }

    #[test]
    fn test_violations_ordered_by_cell() {
        let mut cells = HashMap::new();
        for row in 0..4 {
            for col in 0..4 {
                cells.insert(
                    (row, col),
                    Cell {
                        formula: Some(<Box<str>>::from("=Sheet1!A1+Sheet2!B2+Sheet3!C3")),
                        is_array: false,
                        num_fmt: None,
                        row,
                        col,
                        value: CellValue::Empty,
                    },
                );
            }
        }
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            ..Default::default()
        };

        let violations = run_violations(&sheet);

        let cells: Vec<_> = violations
            .iter()
            .map(|v| v.data::<MultipleSheetRefData>().unwrap().cell.clone())
            .collect();
        let mut sorted = cells.clone();
        sorted.sort();
        assert_eq!(cells.len(), 16);
        assert_eq!(cells, sorted);
    }

    #[test]
    fn test_below_threshold_passes() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=Sheet1!A1+Sheet2!B2")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let violations = run_violations(&sheet);

        assert!(violations.is_empty());
    }
}

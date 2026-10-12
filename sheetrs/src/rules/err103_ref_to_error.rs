//! ERR103: Ref to Excel Error detection
//!
//! Description: Flags formulas referencing cells that currently store an error (#REF!, #DIV/0!, etc.).
//!
//! Violations for this rule are emitted by [`super::err102_error_cells::ErrorCellsRule::on_workbook_end`]
//! because the classification requires the error set, dependency graph, and circular-cell set
//! in one place. This struct is registered as a walker rule only for configuration enable/disable.

use super::WalkerRule;
use crate::violation::{CellReference, FormatContext, RuleId, ViolationData};

/// Rule that identifies formulas referencing cells with errors.
///
/// This walker rule is a no-op. Actual violations are emitted by ERR102's
/// `on_workbook_end` to avoid duplicate graph traversal.
pub struct RefToErrorRule;

/// Incident data for ERR103 (walker path).
#[derive(Debug)]
pub struct RefToErrorData {
    /// Location of the original error cell in the dependency graph.
    pub error_source: (u16, u32, u32),
}

impl ViolationData for RefToErrorData {
    fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        let (idx, r, c) = self.error_source;
        let sheet_name = ctx.workbook.sheet_name_by_index(idx).unwrap_or("Unknown");
        let cell_ref = CellReference::new(r, c);
        format!("Formula references cell with error: {sheet_name}!{cell_ref}")
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for RefToErrorRule {
    fn id(&self) -> RuleId {
        RuleId::Err103
    }

    fn name(&self) -> &'static str {
        "Ref to Excel Error"
    }

    fn category(&self) -> super::RuleCategory {
        super::RuleCategory::ExcelErrors
    }

    // All hooks are no-op: violations are emitted by ERR102's on_workbook_end.
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Workbook};
    use crate::rules::LinterContext;
    use crate::rules::err102_error_cells::ErrorCellsRule;
    use crate::violation::ExcelError;
    use std::collections::HashMap;

    #[test]
    fn test_ref_to_error_violation_data() {
        let data = RefToErrorData {
            error_source: (0, 2, 3),
        };

        let sheets = vec![crate::reader::workbook::Sheet {
            name: "Data".to_string(),
            sheet_index: 0,
            cells: HashMap::new(),
            ..Default::default()
        }];
        let workbook = Workbook {
            sheets,
            ..Default::default()
        };
        let ctx = FormatContext {
            workbook: &workbook,
        };

        let msg = data.format_message(&ctx);
        assert!(
            msg.contains("Data!D3"),
            "Message should contain 'Data!D3', got: {msg}"
        );
    }

    #[test]
    fn test_err103_multiple_error_sources_uses_first() {
        // A1 error, C1 error, B1 formula depends on both → ERR103 uses first dep
        let sheet = crate::reader::workbook::Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: {
                let mut m = HashMap::new();
                m.insert(
                    (0, 0),
                    Cell {
                        row: 0,
                        col: 0,
                        value: CellValue::Error(ExcelError::Ref),
                        formula: None,
                        is_array: false,
                        num_fmt: None,
                    },
                );
                m.insert(
                    (0, 2),
                    Cell {
                        row: 0,
                        col: 2,
                        value: CellValue::Error(ExcelError::NA),
                        formula: None,
                        is_array: false,
                        num_fmt: None,
                    },
                );
                m.insert(
                    (0, 1),
                    Cell {
                        row: 0,
                        col: 1,
                        value: CellValue::Number(0.0),
                        formula: Some(Box::from("A1+C1")),
                        is_array: false,
                        num_fmt: None,
                    },
                );
                m
            },
            ..Default::default()
        };

        let rule = ErrorCellsRule;
        let mut ctx = LinterContext::default();

        for cell in sheet.cells.values() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }

        // B1 depends on A1 and C1 (A1 listed first)
        ctx.cell_dependencies
            .insert((0, 0, 1), vec![(0, 0, 0), (0, 0, 2)]);

        let workbook = Workbook::default();
        let violations = rule.on_workbook_end(&workbook, &mut ctx);

        // ERR103 on B1, ERR102 on C1 (A1 consumed by err103), possibly ERR102 on C1 too
        let err103_violations: Vec<_> = violations
            .iter()
            .filter(|v| v.rule_id == RuleId::Err103)
            .collect();
        assert_eq!(err103_violations.len(), 1, "Should have exactly 1 ERR103");

        // Verify the error source is A1 (first dep)
        let data = err103_violations[0]
            .data::<RefToErrorData>()
            .expect("Should have RefToErrorData");
        assert_eq!(data.error_source, (0, 0, 0), "Error source should be A1");
    }
}

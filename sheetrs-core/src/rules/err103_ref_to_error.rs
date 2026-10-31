//! ERR103: Ref to Excel Error
//!
//! A formula that reads a cell holding an error returns an error too, and one broken input can
//! take down a whole chain of results. Pointing at the formula that consumes the error tells the
//! reader which calculation is affected and where to look for the cause.
//!
//! # Reports
//!
//! One violation per formula cell (unit: the cell) with a reference to a cell that holds an error
//! (see [`err102_error_cells`](super::err102_error_cells) for what counts as an error cell). When
//! a formula reads several error cells, the message names the first one in the formula's reference
//! order, as `Sheet!Cell`. The error cell that is named is not also reported as ERR102.
//!
//! The violations are produced by the ERR102 pass at the end of the workbook, because the
//! classification needs the error cells, the dependency graph and the circular-cell set together.
//! The rule type itself does nothing; it exists so that ERR103 can be enabled or disabled on its
//! own. Enabling it runs ERR102 and CALC202 as well, and their violations are only shown if they
//! are enabled.
//!
//! # Does not report
//!
//! - A formula that holds an error but reads none: that is the source, ERR102.
//! - Formulas, or error cells, on an active circular reference path: CALC202 reports the loop.
//! - Ranges read through their interior only: references are tracked as in CALC202, a range by its
//!   two corner cells, so an error in the middle of `A1:A10` is not seen by `SUM(A1:A10)`.
//! - Whole column or row references, which the dependency graph does not track (REF307).
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always `Error`.

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
    use crate::rules::collect_violations;
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
            cells: Default::default(),
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
                        value: CellValue::Error(ExcelError::Ref),
                        formula: None,
                        is_array: false,
                        num_fmt: None,
                    },
                );
                m.insert(
                    (0, 2),
                    Cell {
                        value: CellValue::Error(ExcelError::NA),
                        formula: None,
                        is_array: false,
                        num_fmt: None,
                    },
                );
                m.insert(
                    (0, 1),
                    Cell {
                        value: CellValue::Number(0.0),
                        formula: Some(Box::from("A1+C1")),
                        is_array: false,
                        num_fmt: None,
                    },
                );
                m.into()
            },
            ..Default::default()
        };

        let mut rule = ErrorCellsRule::default();
        let mut ctx = LinterContext::default();

        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }

        // B1 depends on A1 and C1 (A1 listed first)
        ctx.cell_dependencies
            .insert((0, 0, 1), vec![(0, 0, 0), (0, 0, 2)]);

        let workbook = Workbook::default();
        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));

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

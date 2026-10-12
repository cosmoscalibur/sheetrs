//! ERR102: Error cells detection
//!
//! Description: Identifies cells containing raw Excel calculation error codes (#REF!, #DIV/0!, etc.).
//!
//! This rule operates as a walker collecting error cells during `on_cell`, then in
//! `on_workbook_end` classifies each into ERR102 (standalone error), ERR103
//! (formula referencing an error cell), or suppressed (CALC202 circular dominance).

use super::{LinterContext, WalkerRule};
use crate::reader::{Cell, Sheet, Workbook};
use crate::rules::err103_ref_to_error::RefToErrorData;
use crate::violation::{
    CellReference, ExcelError, FormatContext, RuleId, Severity, Violation, ViolationData,
    ViolationScope,
};
use std::collections::HashSet;

/// Rule that identifies cells containing error values.
///
/// During the cell walk it records error locations into `LinterContext::error_cells`.
/// After all cells are visited, `on_workbook_end` cross-references the error set
/// with the dependency graph and circular-cell set to emit ERR102 or ERR103.
pub struct ErrorCellsRule;

/// Incident data for ERR102 (walker path).
#[derive(Debug)]
pub struct ErrorCellData {
    /// The Excel error type found in the cell.
    pub error: ExcelError,
}

impl ViolationData for ErrorCellData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!("Cell contains error value: {}", self.error)
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for ErrorCellsRule {
    fn id(&self) -> RuleId {
        RuleId::Err102
    }

    fn name(&self) -> &'static str {
        "Excel Error"
    }

    fn category(&self) -> super::RuleCategory {
        super::RuleCategory::ExcelErrors
    }

    fn on_cell(&mut self, sheet: &Sheet, cell: &Cell, ctx: &mut LinterContext) -> Vec<Violation> {
        let loc = (sheet.sheet_index, cell.row, cell.col);

        if let Some(error) = cell.value.as_error() {
            ctx.error_cells.insert(loc, error);
        } else if let Some(formula) = cell.as_formula() {
            // Scan formula text for error literals (covers ODS formulas with embedded errors)
            for &literal in ExcelError::ERROR_LITERALS {
                if formula.contains(literal) {
                    if let Some(err) = ExcelError::from_cell_str(literal) {
                        ctx.error_cells.insert(loc, err);
                    }
                    break;
                }
            }
        }

        Vec::new()
    }

    fn on_workbook_end(&mut self, _workbook: &Workbook, ctx: &mut LinterContext) -> Vec<Violation> {
        let mut violations = Vec::new();

        // Collect error cells NOT on a circular reference path (CALC202 dominance).
        let non_circular_errors: Vec<_> = ctx
            .error_cells
            .iter()
            .filter(|(loc, _)| !ctx.circular_cells.contains(loc))
            .map(|(loc, err)| (*loc, *err))
            .collect();

        // Build a set of error cell locations for fast lookup.
        let error_locs: HashSet<(u16, u32, u32)> =
            non_circular_errors.iter().map(|(loc, _)| *loc).collect();

        // Track which error cells are "consumed" by ERR103 (referenced by a formula).
        let mut consumed_by_err103 = HashSet::new();

        // For each formula cell in the dependency graph, check if any of its
        // dependencies point to a non-circular error cell → ERR103.
        for (formula_loc, deps) in &ctx.cell_dependencies {
            // Skip formula cells that are themselves on a circular path.
            if ctx.circular_cells.contains(formula_loc) {
                continue;
            }

            // Find the first dependency that is an error cell.
            let first_error_dep = deps.iter().find(|dep| error_locs.contains(dep));

            if let Some(error_source) = first_error_dep {
                let (sheet_index, row, col) = formula_loc;
                violations.push(Violation::with_data(
                    RuleId::Err103,
                    ViolationScope::Cell(*sheet_index, CellReference::new(*row, *col)),
                    RefToErrorData {
                        error_source: *error_source,
                    },
                    Severity::Error,
                ));
                consumed_by_err103.insert(*error_source);
            }
        }

        // Remaining non-circular error cells that were not consumed by ERR103 → ERR102.
        for (loc, error) in &non_circular_errors {
            if consumed_by_err103.contains(loc) {
                continue;
            }
            let (sheet_index, row, col) = loc;
            violations.push(Violation::with_data(
                RuleId::Err102,
                ViolationScope::Cell(*sheet_index, CellReference::new(*row, *col)),
                ErrorCellData { error: *error },
                Severity::Error,
            ));
        }

        violations
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::LinterContext;
    use std::collections::HashMap;

    fn make_cell(row: u32, col: u32, value: CellValue, formula: Option<&str>) -> Cell {
        Cell {
            row,
            col,
            value,
            formula: formula.map(Box::from),
            is_array: false,
            num_fmt: None,
        }
    }

    fn make_sheet(name: &str, index: u16, cells: Vec<Cell>) -> Sheet {
        let mut cell_map = HashMap::new();
        for c in cells {
            cell_map.insert((c.row, c.col), c);
        }
        Sheet {
            name: name.to_string(),
            sheet_index: index,
            cells: cell_map,
            ..Default::default()
        }
    }

    #[test]
    fn test_err102_standalone_error() {
        let sheet = make_sheet(
            "Sheet1",
            0,
            vec![
                make_cell(0, 0, CellValue::Error(ExcelError::DivZero), None),
                make_cell(1, 0, CellValue::Number(42.0), None),
            ],
        );

        let mut rule = ErrorCellsRule;
        let mut ctx = LinterContext::default();

        // Walk cells
        for cell in sheet.cells.values() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }

        // Workbook end — no circular cells, no deps
        let workbook = Workbook::default();
        let violations = rule.on_workbook_end(&workbook, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Err102);
        assert!(violations[0].message().contains("#DIV/0!"));
    }

    #[test]
    fn test_err102_formula_literal_error() {
        let sheet = make_sheet(
            "Sheet1",
            0,
            vec![make_cell(0, 0, CellValue::Empty, Some("SUM(A1, [#REF!])"))],
        );

        let mut rule = ErrorCellsRule;
        let mut ctx = LinterContext::default();

        for cell in sheet.cells.values() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }

        let workbook = Workbook::default();
        let violations = rule.on_workbook_end(&workbook, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Err102);
        assert!(violations[0].message().contains("#REF!"));
    }

    #[test]
    fn test_err103_ref_to_error() {
        // A1 has error, B1 formula =A1 depends on A1 → ERR103 on B1
        let sheet = make_sheet(
            "Sheet1",
            0,
            vec![
                make_cell(0, 0, CellValue::Error(ExcelError::Ref), None),
                make_cell(0, 1, CellValue::Number(0.0), Some("A1")),
            ],
        );

        let mut rule = ErrorCellsRule;
        let mut ctx = LinterContext::default();

        for cell in sheet.cells.values() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }

        // Simulate dependency: B1 -> A1
        ctx.cell_dependencies.insert((0, 0, 1), vec![(0, 0, 0)]);

        let workbook = Workbook::default();
        let violations = rule.on_workbook_end(&workbook, &mut ctx);

        // Should have ERR103 on B1 (not ERR102 on A1, since A1 is consumed)
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Err103);
    }

    #[test]
    fn test_calc202_dominance() {
        // A1 has error + is on circular path → no ERR102/ERR103
        let sheet = make_sheet(
            "Sheet1",
            0,
            vec![make_cell(
                0,
                0,
                CellValue::Error(ExcelError::Ref),
                Some("A1"),
            )],
        );

        let mut rule = ErrorCellsRule;
        let mut ctx = LinterContext::default();

        for cell in sheet.cells.values() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }

        // Mark as circular (populated by CALC202)
        ctx.circular_cells.insert((0, 0, 0));

        let workbook = Workbook::default();
        let violations = rule.on_workbook_end(&workbook, &mut ctx);

        assert!(
            violations.is_empty(),
            "CALC202 dominance should suppress ERR102"
        );
    }
}

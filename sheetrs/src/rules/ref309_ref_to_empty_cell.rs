//! REF309: Reference to Empty Cell detection
//!
//! Description: Flags formulas referencing cells that contain no data or formulas.
//! Suppresses on sheets where all cells are empty (REF303 territory).

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::Workbook;
use crate::reader::workbook::CellValue;
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::HashSet;

/// Rule that identifies formulas referencing empty cells.
///
/// **Dependency:** requires CALC202 (Circular References) to be active,
/// as it populates `ctx.cell_dependencies`.
pub struct RefToEmptyCellRule;

/// Incident data for REF309.
#[derive(Debug)]
pub struct RefToEmptyCellData {
    /// Location of the empty cell being referenced.
    pub empty_cell: (u16, u32, u32),
}

impl ViolationData for RefToEmptyCellData {
    fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        let (sheet_idx, row, col) = self.empty_cell;
        let sheet_name = ctx
            .workbook
            .sheets
            .get(sheet_idx as usize)
            .map(|s| s.name.as_str())
            .unwrap_or("?");
        format!(
            "Formula references empty cell: {}!{}",
            sheet_name,
            CellReference::new(row, col)
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for RefToEmptyCellRule {
    fn id(&self) -> RuleId {
        RuleId::Ref309
    }

    fn name(&self) -> &'static str {
        "Ref to Empty Cell"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Reference
    }

    fn on_workbook_end(
        &mut self,
        workbook: &Workbook,
        ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        if ctx.cell_dependencies.is_empty() {
            return;
        }

        // Build set of non-empty cell positions per sheet.
        // A cell is "non-empty" if it has a formula or a non-empty value.
        let mut non_empty: HashSet<(u16, u32, u32)> = HashSet::new();
        let mut sheets_with_content: HashSet<u16> = HashSet::new();

        for sheet in &workbook.sheets {
            for cell in sheet.cells.values() {
                let has_content = cell.is_formula() || !matches!(cell.value, CellValue::Empty);
                if has_content {
                    non_empty.insert((sheet.sheet_index, cell.row, cell.col));
                    sheets_with_content.insert(sheet.sheet_index);
                }
            }
        }

        for (formula_loc, deps) in &ctx.cell_dependencies {
            // Suppress on sheets with no content at all (REF303 territory)
            if !sheets_with_content.contains(&formula_loc.0) {
                continue;
            }

            // Find first empty dependency
            if let Some(empty_dep) = deps.iter().find(|dep| !non_empty.contains(dep)) {
                out.push(Violation::with_data(
                    RuleId::Ref309,
                    ViolationScope::Cell(
                        formula_loc.0,
                        CellReference::new(formula_loc.1, formula_loc.2),
                    ),
                    RefToEmptyCellData {
                        empty_cell: *empty_dep,
                    },
                    Severity::Info,
                ));
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet, Workbook};
    use crate::rules::LinterContext;
    use crate::rules::collect_violations;
    use std::collections::HashMap;

    fn make_number_cell(row: u32, col: u32, n: f64) -> Cell {
        Cell {
            row,
            col,
            value: CellValue::Number(n),
            formula: None,
            is_array: false,
            num_fmt: None,
        }
    }

    fn make_formula_cell(row: u32, col: u32, formula: &str) -> Cell {
        Cell {
            row,
            col,
            value: CellValue::Empty,
            formula: Some(Box::from(formula)),
            is_array: false,
            num_fmt: None,
        }
    }

    fn make_sheet_with_index(index: u16, name: &str, cells: Vec<Cell>) -> Sheet {
        let mut cell_map = HashMap::new();
        for cell in cells {
            cell_map.insert((cell.row, cell.col), cell);
        }
        Sheet {
            name: name.to_string(),
            sheet_index: index,
            cells: cell_map,
            ..Default::default()
        }
    }

    fn make_workbook(sheets: Vec<Sheet>) -> Workbook {
        Workbook {
            sheets,
            ..Default::default()
        }
    }

    #[test]
    fn test_formula_refs_empty_cell() {
        // A1 is empty (not in cells), B1 = =A1
        let sheet = make_sheet_with_index(0, "Sheet1", vec![make_formula_cell(0, 1, "A1")]);
        let workbook = make_workbook(vec![sheet]);

        let mut rule = RefToEmptyCellRule;
        let mut ctx = LinterContext::default();
        // B1(0,1) depends on A1(0,0) which is not in the sheet
        ctx.cell_dependencies.insert((0, 0, 1), vec![(0, 0, 0)]);

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_formula_refs_populated_cell() {
        let sheet = make_sheet_with_index(
            0,
            "Sheet1",
            vec![make_number_cell(0, 0, 42.0), make_formula_cell(0, 1, "A1")],
        );
        let workbook = make_workbook(vec![sheet]);

        let mut rule = RefToEmptyCellRule;
        let mut ctx = LinterContext::default();
        ctx.cell_dependencies.insert((0, 0, 1), vec![(0, 0, 0)]);

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_cross_sheet_ref_to_empty() {
        let sheet1 = make_sheet_with_index(0, "Sheet1", vec![make_formula_cell(0, 1, "Sheet2!A1")]);
        let sheet2 = make_sheet_with_index(1, "Sheet2", vec![]);
        let workbook = make_workbook(vec![sheet1, sheet2]);

        let mut rule = RefToEmptyCellRule;
        let mut ctx = LinterContext::default();
        ctx.cell_dependencies.insert((0, 0, 1), vec![(1, 0, 0)]); // Sheet1!B1 → Sheet2!A1

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_empty_sheet_suppression() {
        // Formula on a sheet with NO content → suppress (REF303 territory)
        let sheet = make_sheet_with_index(0, "Sheet1", vec![]);
        let workbook = make_workbook(vec![sheet]);

        let mut rule = RefToEmptyCellRule;
        let mut ctx = LinterContext::default();
        // Formula at A1 depends on B1 — but A1 isn't even in the sheet
        ctx.cell_dependencies.insert((0, 0, 0), vec![(0, 0, 1)]);

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 0, "Should suppress on empty sheets");
    }

    #[test]
    fn test_multiple_empty_deps_single_violation() {
        let sheet = make_sheet_with_index(0, "Sheet1", vec![make_formula_cell(0, 2, "A1+A2")]);
        let workbook = make_workbook(vec![sheet]);

        let mut rule = RefToEmptyCellRule;
        let mut ctx = LinterContext::default();
        ctx.cell_dependencies
            .insert((0, 0, 2), vec![(0, 0, 0), (0, 1, 0)]);

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 1, "First empty dep only");
    }

    #[test]
    fn test_cell_with_formula_not_empty() {
        let sheet = make_sheet_with_index(
            0,
            "Sheet1",
            vec![
                make_formula_cell(0, 0, "NOW()"),
                make_formula_cell(0, 1, "A1"),
            ],
        );
        let workbook = make_workbook(vec![sheet]);

        let mut rule = RefToEmptyCellRule;
        let mut ctx = LinterContext::default();
        ctx.cell_dependencies.insert((0, 0, 1), vec![(0, 0, 0)]);

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 0, "A cell with a formula is not empty");
    }
}

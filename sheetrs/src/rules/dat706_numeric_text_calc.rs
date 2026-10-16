//! DAT706: Numeric Text in Calculation
//!
//! Description: Flags formulas that reference cells containing text-typed values
//! that look numeric (e.g., a cell storing "10" as text used in =A1+1).

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::workbook::CellValue;
use crate::reader::{Cell, Sheet, Workbook};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};

/// Rule that identifies numeric text cells referenced in calculations.
///
/// **Dependency:** requires CALC202 (Circular References) to be active,
/// as it populates `ctx.cell_dependencies`.
pub struct NumericTextCalcRule;

/// Incident data for DAT706.
#[derive(Debug)]
pub struct NumericTextCalcData {
    /// Location of the first numeric-text dependency.
    pub text_cell: (u16, u32, u32),
}

impl ViolationData for NumericTextCalcData {
    fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        let (sheet_idx, row, col) = self.text_cell;
        let sheet_name = ctx
            .workbook
            .sheets
            .get(sheet_idx as usize)
            .map(|s| s.name.as_str())
            .unwrap_or("?");
        format!(
            "Formula references numeric text cell: {}!{}. Convert text to number.",
            sheet_name,
            CellReference::new(row, col)
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for NumericTextCalcRule {
    fn id(&self) -> RuleId {
        RuleId::Data706
    }

    fn name(&self) -> &'static str {
        "Numeric Text Calc"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Data
    }

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        cell: &Cell,
        ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        // Collect numeric-text cells: text values that parse as f64
        if let CellValue::Text(ref t) = cell.value
            && !cell.is_formula()
            && t.trim().parse::<f64>().is_ok()
        {
            ctx.numeric_text_cells
                .insert((sheet.sheet_index, cell.row, cell.col));
        }
    }

    fn on_workbook_end(
        &mut self,
        _workbook: &Workbook,
        ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        if ctx.numeric_text_cells.is_empty() || ctx.cell_dependencies.is_empty() {
            return;
        }

        for (formula_loc, deps) in &ctx.cell_dependencies {
            // Find the first dependency that is a numeric-text cell
            if let Some(text_dep) = deps.iter().find(|dep| ctx.numeric_text_cells.contains(dep)) {
                out.push(Violation::with_data(
                    RuleId::Data706,
                    ViolationScope::Cell(
                        formula_loc.0,
                        CellReference::new(formula_loc.1, formula_loc.2),
                    ),
                    NumericTextCalcData {
                        text_cell: *text_dep,
                    },
                    Severity::Warning,
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
    use std::sync::Arc;

    fn make_text_cell(row: u32, col: u32, text: &str) -> Cell {
        Cell {
            row,
            col,
            value: CellValue::Text(Arc::from(text)),
            formula: None,
            is_array: false,
            num_fmt: None,
        }
    }

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

    fn make_workbook(sheets: Vec<Sheet>) -> Workbook {
        Workbook {
            sheets,
            ..Default::default()
        }
    }

    #[test]
    fn test_text_number_referenced_by_formula() {
        // A1 = Text("42"), B1 = =A1+1
        let sheet = make_sheet(vec![
            make_text_cell(0, 0, "42"),
            make_formula_cell(0, 1, "A1+1"),
        ]);
        let workbook = make_workbook(vec![sheet.clone()]);

        let mut rule = NumericTextCalcRule;
        let mut ctx = LinterContext::default();

        // Simulate walk
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx, &mut Vec::new());
        }
        // Simulate CALC202 populating deps: B1(0,1) depends on A1(0,0)
        ctx.cell_dependencies.insert((0, 0, 1), vec![(0, 0, 0)]);

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_actual_number_no_violation() {
        let sheet = make_sheet(vec![
            make_number_cell(0, 0, 42.0),
            make_formula_cell(0, 1, "A1+1"),
        ]);
        let workbook = make_workbook(vec![sheet.clone()]);

        let mut rule = NumericTextCalcRule;
        let mut ctx = LinterContext::default();

        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx, &mut Vec::new());
        }
        ctx.cell_dependencies.insert((0, 0, 1), vec![(0, 0, 0)]);

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_non_numeric_text_no_violation() {
        let sheet = make_sheet(vec![
            make_text_cell(0, 0, "hello"),
            make_formula_cell(0, 1, "A1+1"),
        ]);
        let workbook = make_workbook(vec![sheet.clone()]);

        let mut rule = NumericTextCalcRule;
        let mut ctx = LinterContext::default();

        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx, &mut Vec::new());
        }
        ctx.cell_dependencies.insert((0, 0, 1), vec![(0, 0, 0)]);

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_no_formula_references_text() {
        // A1 = Text("42"), B1 = Number(1) (no formula)
        let sheet = make_sheet(vec![
            make_text_cell(0, 0, "42"),
            make_number_cell(0, 1, 1.0),
        ]);
        let workbook = make_workbook(vec![sheet.clone()]);

        let mut rule = NumericTextCalcRule;
        let mut ctx = LinterContext::default();

        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx, &mut Vec::new());
        }
        // No dependencies (no formulas)

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_multiple_text_deps() {
        let sheet = make_sheet(vec![
            make_text_cell(0, 0, "1"),
            make_text_cell(1, 0, "2"),
            make_formula_cell(0, 1, "A1+A2"),
        ]);
        let workbook = make_workbook(vec![sheet.clone()]);

        let mut rule = NumericTextCalcRule;
        let mut ctx = LinterContext::default();

        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx, &mut Vec::new());
        }
        ctx.cell_dependencies
            .insert((0, 0, 1), vec![(0, 0, 0), (0, 1, 0)]);

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 1, "Only one violation per formula cell");
    }

    #[test]
    fn test_empty_string_not_numeric() {
        let sheet = make_sheet(vec![
            make_text_cell(0, 0, ""),
            make_formula_cell(0, 1, "A1+1"),
        ]);
        let workbook = make_workbook(vec![sheet.clone()]);

        let mut rule = NumericTextCalcRule;
        let mut ctx = LinterContext::default();

        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx, &mut Vec::new());
        }
        ctx.cell_dependencies.insert((0, 0, 1), vec![(0, 0, 0)]);

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 0);
    }
}

//! DATA706: Numeric text used in a calculation
//!
//! A cell that holds `"10"` as text looks like the number 10 but is not one: `SUM` and `COUNT` skip
//! it, `+` converts it and `=` comparisons treat it as a different type. A formula that reads such a
//! cell can therefore return a result that depends on the function used, without any visible error.
//!
//! # Reports
//!
//! One violation per formula cell (unit: the formula cell), reported once however many numeric-text
//! cells it reads; the message names the first one found. A numeric-text cell is a constant (not a
//! formula) whose text, trimmed of white space, parses as a floating-point number (`42`, `-1.5`,
//! `1e3`; also the forms Rust accepts such as `inf`). A formula reads such a cell when the cell is
//! among its references. Any reference counts, whatever the function: the rule does not check that the
//! reference is used arithmetically. A range counts only through its two corner cells.
//!
//! # Does not report
//!
//! - Text that does not parse as a number, and empty text.
//! - Text returned by a formula (`=IF(A1,"0","1")`): a result, not stored data, as in DATA702.
//! - Numeric text inside a range, away from its corners, and references to whole columns or rows:
//!   the dependency graph does not track them (REF307 flags whole-column references).
//! - The numeric-text cell itself: that is DATA702, which reports numbers stored as text. This rule
//!   reports where they are used.
//!
//! # Parameters
//!
//! None. The rule needs CALC202 enabled: it reads the dependency graph that CALC202 builds, and
//! reports nothing when that graph is empty.
//!
//! # Severity
//!
//! Always `Warning`.

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
        at: (u32, u32),
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
                .insert((sheet.sheet_index, at.0, at.1));
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

    fn make_text_cell(row: u32, col: u32, text: &str) -> ((u32, u32), Cell) {
        (
            (row, col),
            Cell {
                value: CellValue::Text(Arc::from(text)),
                formula: None,
                is_array: false,
                num_fmt: None,
            },
        )
    }

    fn make_number_cell(row: u32, col: u32, n: f64) -> ((u32, u32), Cell) {
        (
            (row, col),
            Cell {
                value: CellValue::Number(n),
                formula: None,
                is_array: false,
                num_fmt: None,
            },
        )
    }

    fn make_formula_cell(row: u32, col: u32, formula: &str) -> ((u32, u32), Cell) {
        (
            (row, col),
            Cell {
                value: CellValue::Empty,
                formula: Some(Box::from(formula)),
                is_array: false,
                num_fmt: None,
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
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
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

        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
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

        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
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

        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
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

        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
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

        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        ctx.cell_dependencies.insert((0, 0, 1), vec![(0, 0, 0)]);

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 0);
    }
}

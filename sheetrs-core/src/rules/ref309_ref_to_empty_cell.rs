//! REF309: Reference to empty cell
//!
//! A formula that reads a cell holding neither a value nor a formula. The result silently treats
//! the empty cell as 0 or as an empty string, which usually means a mistyped address, a deleted
//! input or a reference shifted off its data.
//!
//! # Reports
//!
//! One violation per formula cell (the unit is the cell) that has at least one empty dependency;
//! the message names the first empty dependency found, even when there are several. A dependency
//! is empty when the referenced cell does not exist or has an empty value and no formula (a
//! formula with an empty saved value is not empty; a text value, even `""`, is not empty).
//!
//! The dependencies are those of [`LinterContext::cell_dependencies`], filled by CALC202
//! ([`super::calc202_circular_references`], which must be active; without it nothing is
//! reported). They come from the parsed formula, so string literals are not references and
//! references to unknown sheets are ignored. A range is represented by its two corner cells only:
//! empty cells inside a range are not seen, which is intended since ranges commonly contain
//! blanks. If a formula cannot be parsed, CALC202 falls back to a textual scan and every
//! reference-like text counts.
//!
//! # Does not report
//!
//! - Formulas on a sheet that has no content at all (no formula and no value): an empty sheet is
//!   REF303.
//! - Whole-column and whole-row references (not recorded as dependencies; REF307).
//! - Interior cells of a range, as above.
//! - A reference to an empty cell that is part of a loop or an error: CALC202, ERR103.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always info.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::workbook::CellValue;
use crate::reader::{Sheet, Workbook};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::{HashMap, HashSet};

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

        // A cell is "non-empty" if it has a formula or a non-empty value. The sheets' own cells
        // answer that; a copy of every position (a set of all the cells of the workbook) cost
        // about 40 bytes for each.
        let has_content = |cell: &crate::reader::Cell| {
            cell.is_formula() || !matches!(cell.value, CellValue::Empty)
        };
        let sheets: HashMap<u16, &Sheet> =
            workbook.sheets.iter().map(|s| (s.sheet_index, s)).collect();
        let sheets_with_content: HashSet<u16> = workbook
            .sheets
            .iter()
            .filter(|sheet| sheet.cells.values().any(has_content))
            .map(|sheet| sheet.sheet_index)
            .collect();
        let non_empty = |at: &(u16, u32, u32)| {
            sheets
                .get(&at.0)
                .and_then(|sheet| sheet.cells.get(&(at.1, at.2)))
                .is_some_and(has_content)
        };

        for (formula_loc, deps) in &ctx.cell_dependencies {
            // Suppress on sheets with no content at all (REF303 territory)
            if !sheets_with_content.contains(&formula_loc.0) {
                continue;
            }

            // Find first empty dependency
            if let Some(empty_dep) = deps.iter().find(|dep| !non_empty(dep)) {
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

    fn make_sheet_with_index(index: u16, name: &str, cells: Vec<((u32, u32), Cell)>) -> Sheet {
        let mut cell_map = HashMap::new();
        for (at, cell) in cells {
            cell_map.insert(at, cell);
        }
        Sheet {
            name: name.to_string(),
            sheet_index: index,
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

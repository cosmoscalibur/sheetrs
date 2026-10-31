//! ERR102: Error Cells
//!
//! A cell that shows an error value (`#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A` ...) is a result that
//! cannot be used, and every formula that reads it fails in turn. The rule reports where the error
//! starts rather than every cell it spreads to.
//!
//! # Reports
//!
//! One violation per error cell (unit: the cell) that is the source of an error: a cell whose saved
//! value is an error, or whose formula text contains an error literal (`#NULL!`, `#DIV/0!`,
//! `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, `#SPILL!`, `#CALC!`; this also catches ODS
//! formulas that embed one). The message names the error. Walking all cells and classifying them is
//! done here at the end of the workbook, for this rule and for ERR103.
//!
//! # Does not report
//!
//! - An error cell that a formula reads: that formula is reported by ERR103 instead, and the cell
//!   it reads is not repeated here. In a chain `A1` (error), `B1 = A1`, `C1 = B1`, ERR103 reports
//!   `B1` and `C1`, and this rule only `C1`, the end of the chain.
//! - Error cells on a circular reference path that the saved values show as active: CALC202
//!   reports the loop and takes precedence, so the same cells are not reported twice.
//! - A saved `#NAME?` whose formula calls a function written with the `_xlfn.` (or `_xlws.`)
//!   prefix, outside its quoted text: that prefix marks a function newer than the file format, and
//!   the error only shows that the application that saved the file did not know it, so the cache
//!   does not prove a mistake. A cell that reads such a cell and holds `#NAME?` itself is excluded
//!   too. Exception: `_xlfn._LONGTEXT(…)` is not an Excel function (LibreOffice 26.2 wrote a text
//!   literal of more than 255 characters as pieces joined by it), so its `#NAME?` is reported
//!   unless the `LONGTEXT` feature is enabled (see Parameters).
//! - Errors that are not stored in the file: nothing is recalculated, so a formula whose saved
//!   value is not an error is not evaluated to find one.
//!
//! The dependency graph used to match readers to error cells is filled by CALC202, which therefore
//! runs whenever this rule does, even if it is not itself reported.
//!
//! # Parameters
//!
//! - `enable_features` (array of strings, default empty; global only): compatibility features. With
//!   `"LONGTEXT"` a `#NAME?` of a formula that calls `_xlfn._LONGTEXT(…)` is not an error, the same
//!   feature `sheetc` has to join those pieces when the program is built. A name that is not known
//!   makes the configuration invalid.
//!
//! # Severity
//!
//! Always `Error`.

use super::helpers::future_calls;
use super::{CellDependencyMap, LinterContext, WalkerRule};
use crate::config::LinterConfig;
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
#[derive(Default)]
pub struct ErrorCellsRule {
    /// The `LONGTEXT` compatibility feature is on: `_xlfn._LONGTEXT(…)` is accepted.
    longtext: bool,
}

impl ErrorCellsRule {
    /// Create the rule, reading `enable_features` from the configuration.
    pub fn new(config: &LinterConfig) -> Self {
        Self {
            longtext: config.feature_enabled("LONGTEXT"),
        }
    }

    /// Whether a saved `#NAME?` of a cell with this formula comes from the application that saved
    /// the file not knowing a function written with the `_xlfn.` prefix, and not from a mistake of
    /// the workbook. A function that is not an Excel one (`_xlfn._LONGTEXT`) is a mistake unless
    /// the `LONGTEXT` feature is on.
    fn is_unknown_to_the_saver(&self, formula: &str) -> bool {
        let calls = future_calls(formula);
        calls.any && (!calls.longtext || self.longtext)
    }
}

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

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        at: (u32, u32),
        cell: &Cell,
        ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        let loc = (sheet.sheet_index, at.0, at.1);

        if let Some(error) = cell.value.as_error() {
            if error == ExcelError::Name
                && cell
                    .as_formula()
                    .is_some_and(|f| self.is_unknown_to_the_saver(f))
            {
                ctx.compat_name_cells.insert(loc);
            } else {
                ctx.error_cells.insert(loc, error);
            }
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
    }

    fn on_workbook_end(
        &mut self,
        _workbook: &Workbook,
        ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        // A `#NAME?` that a cell inherits from a compatibility cell is not an error either: the
        // readers of each such cell are visited once, from the compatibility cells outwards.
        if !ctx.compat_name_cells.is_empty() {
            let mut readers = CellDependencyMap::new();
            for (loc, err) in &ctx.error_cells {
                if *err == ExcelError::Name
                    && let Some(deps) = ctx.cell_dependencies.get(loc)
                {
                    for dep in deps {
                        readers.entry(*dep).or_default().push(*loc);
                    }
                }
            }
            let mut pending: Vec<_> = ctx.compat_name_cells.iter().copied().collect();
            while let Some(source) = pending.pop() {
                for loc in readers.remove(&source).unwrap_or_default() {
                    if ctx.error_cells.remove(&loc).is_some() {
                        ctx.compat_name_cells.insert(loc);
                        pending.push(loc);
                    }
                }
            }
        }

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
                out.push(Violation::with_data(
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
            out.push(Violation::with_data(
                RuleId::Err102,
                ViolationScope::Cell(*sheet_index, CellReference::new(*row, *col)),
                ErrorCellData { error: *error },
                Severity::Error,
            ));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::LinterContext;
    use crate::rules::collect_violations;
    use std::collections::HashMap;

    fn make_cell(
        row: u32,
        col: u32,
        value: CellValue,
        formula: Option<&str>,
    ) -> ((u32, u32), Cell) {
        (
            (row, col),
            Cell {
                value,
                formula: formula.map(Box::from),
                is_array: false,
                num_fmt: None,
            },
        )
    }

    fn make_sheet(name: &str, index: u16, cells: Vec<((u32, u32), Cell)>) -> Sheet {
        let mut cell_map = HashMap::new();
        for (at, c) in cells {
            cell_map.insert(at, c);
        }
        Sheet {
            name: name.to_string(),
            sheet_index: index,
            cells: cell_map.into(),
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

        let mut rule = ErrorCellsRule::default();
        let mut ctx = LinterContext::default();

        // Walk cells
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }

        // Workbook end — no circular cells, no deps
        let workbook = Workbook::default();
        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));

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

        let mut rule = ErrorCellsRule::default();
        let mut ctx = LinterContext::default();

        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }

        let workbook = Workbook::default();
        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));

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

        let mut rule = ErrorCellsRule::default();
        let mut ctx = LinterContext::default();

        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }

        // Simulate dependency: B1 -> A1
        ctx.cell_dependencies.insert((0, 0, 1), vec![(0, 0, 0)]);

        let workbook = Workbook::default();
        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));

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

        let mut rule = ErrorCellsRule::default();
        let mut ctx = LinterContext::default();

        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }

        // Mark as circular (populated by CALC202)
        ctx.circular_cells.insert((0, 0, 0));

        let workbook = Workbook::default();
        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));

        assert!(
            violations.is_empty(),
            "CALC202 dominance should suppress ERR102"
        );
    }

    /// ERR102 violations of a column whose cells hold these formulas with a saved `#NAME?`.
    fn name_errors(formulas: &[&str], config: &LinterConfig) -> usize {
        let cells = formulas
            .iter()
            .enumerate()
            .map(|(row, f)| make_cell(row as u32, 0, CellValue::Error(ExcelError::Name), Some(f)))
            .collect();
        let sheet = make_sheet("Sheet1", 0, cells);
        let mut rule = ErrorCellsRule::new(config);
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let workbook = Workbook::default();
        collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out)).len()
    }

    #[test]
    fn a_saved_name_error_needs_the_formula_to_have_no_prefixed_function() {
        let off = LinterConfig::default();
        // An unknown name is a mistake of the workbook.
        assert_eq!(name_errors(&["FOO(1)"], &off), 1);
        // A function with the `_xlfn.` prefix that the saving application did not know is not.
        assert_eq!(name_errors(&["_xlfn.XLOOKUP(1,B1:B2,C1:C2)"], &off), 0);
        assert_eq!(name_errors(&["_xlfn._xlws.FILTER(B1:B2,C1:C2)"], &off), 0);
        // The prefix inside a string says nothing.
        assert_eq!(name_errors(&["FOO(\"_xlfn.X(\")"], &off), 1);
    }

    #[test]
    fn longtext_is_an_error_unless_the_feature_is_enabled() {
        let formula = "_xlfn._LONGTEXT(\"ab\",\"cd\")";
        assert_eq!(name_errors(&[formula], &LinterConfig::default()), 1);
        let on = LinterConfig::from_toml("[global]\nenable_features = [\"longtext\"]\n").unwrap();
        assert_eq!(name_errors(&[formula], &on), 0);
        // The feature does not excuse another wrapper in the same formula that is not accepted.
        assert_eq!(name_errors(&["FOO(1)"], &on), 1);
    }

    #[test]
    fn a_name_error_inherited_from_such_a_cell_is_not_reported() {
        let sheet = make_sheet(
            "Sheet1",
            0,
            vec![
                make_cell(
                    0,
                    0,
                    CellValue::Error(ExcelError::Name),
                    Some("_xlfn.XLOOKUP(1,C1:C2,D1:D2)"),
                ),
                make_cell(1, 0, CellValue::Error(ExcelError::Name), Some("A1*2")),
                make_cell(2, 0, CellValue::Error(ExcelError::Name), Some("A2+1")),
            ],
        );
        let mut rule = ErrorCellsRule::default();
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        ctx.cell_dependencies.insert((0, 1, 0), vec![(0, 0, 0)]);
        ctx.cell_dependencies.insert((0, 2, 0), vec![(0, 1, 0)]);
        let workbook = Workbook::default();
        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert!(violations.is_empty(), "{violations:?}");
    }

    #[test]
    fn enable_features_accepts_only_known_names() {
        let ok = |toml: &str| {
            LinterConfig::from_toml(toml)
                .unwrap()
                .validate_rules(&std::collections::HashSet::new())
        };
        assert!(ok("[global]\nenable_features = [\"LONGTEXT\"]\n").is_ok());
        assert!(ok("[global]\nenable-features = []\n").is_ok());
        assert!(ok("[global]\nenable_features = [\"NOPE\"]\n").is_err());
        assert!(ok("[global]\nenable_features = \"LONGTEXT\"\n").is_err());
    }
}

//! REF307: Whole-column or whole-row references
//!
//! Formulas that refer to a whole column (`A:A`) or a whole row (`1:1`). Such a reference spans
//! over a million rows or 16 thousand columns: it makes calculation slower, picks up notes and
//! totals added below the data by accident, and is not tracked by the dependency analysis (the
//! cell dependencies of CALC202 leave it out), so other rules do not see it.
//!
//! # Reports
//!
//! The formula cells containing such a reference are grouped per sheet into contiguous blocks
//! (cells that touch by row or column), and one violation is reported per block (the unit is the
//! block, scoped to the sheet), with the bounding range of the block. Whole-column and whole-row
//! cells are grouped separately; a cell that contains both is counted as whole-column only.
//!
//! A whole column (`A:A`, `$B:$D`, `Sheet2!A:A`) or a whole row (`1:1`) is read from the parsed formula,
//! so text between quotes such as `"10:30"` is not a reference. A formula the parser cannot read is
//! searched as text, with its quoted parts masked.
//!
//! # Does not report
//!
//! - Bounded ranges (`A1:A10`) and whole-column references written as `A1:A1048576`.
//! - Performance of bounded but very large ranges, which is out of scope.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always warning.

use super::helpers::{bounding_box, find_contiguous_ranges, mask_quoted, whole_line_refs};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use regex::Regex;
use std::collections::HashMap;
use std::sync::LazyLock;

/// Matches whole-column references (A:A, A:Z, etc.).
static COLUMN_PATTERN: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"\b[A-Z]+:[A-Z]+\b").expect("REF307 column regex must compile"));

/// Matches whole-row references (1:1, 1:100, etc.).
static ROW_PATTERN: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"\b\d+:\d+\b").expect("REF307 row regex must compile"));

/// Per-sheet cell collection: sheet_index → Vec<(row, col, is_column)>.
type SheetCellMap = HashMap<u16, Vec<(u32, u32, bool)>>;

/// Rule that detects whole-column (e.g., A:A) or whole-row (e.g., 1:1) references.
pub struct WholeColumnRowRefsRule {
    /// Per-sheet collected cells.
    sheet_cells: SheetCellMap,
}

impl WholeColumnRowRefsRule {
    /// Create a new instance.
    pub fn new() -> Self {
        Self {
            sheet_cells: HashMap::new(),
        }
    }
}

impl Default for WholeColumnRowRefsRule {
    fn default() -> Self {
        Self::new()
    }
}

/// Incident data for REF307.
#[derive(Debug)]
pub struct WholeRefData {
    /// `true` for whole-column, `false` for whole-row.
    pub is_column: bool,
    /// Bounding box of violating cells (start_row, start_col, end_row, end_col).
    pub range: (u32, u32, u32, u32),
}

impl ViolationData for WholeRefData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let (sr, sc, er, ec) = self.range;
        let range_str = if sr == er && sc == ec {
            CellReference::new(sr, sc).to_string()
        } else {
            format!(
                "{}:{}",
                CellReference::new(sr, sc),
                CellReference::new(er, ec)
            )
        };

        if self.is_column {
            format!(
                "Whole-column reference (e.g., A:A) found in range: {range_str}. Use bounded ranges for better performance."
            )
        } else {
            format!(
                "Whole-row reference (e.g., 1:1) found in range: {range_str}. Use bounded ranges for better performance."
            )
        }
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for WholeColumnRowRefsRule {
    fn id(&self) -> RuleId {
        RuleId::Ref307
    }

    fn name(&self) -> &'static str {
        "Whole Column or Row Reference"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Reference
    }

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        at: (u32, u32),
        cell: &Cell,
        ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        if let Some(formula) = cell.as_formula() {
            // From the parsed formula; one the parser does not read is searched as text, without
            // its quoted parts.
            let (column, row) = whole_line_refs(formula).unwrap_or_else(|| {
                let masked = mask_quoted(ctx.formula_upper(sheet, at, cell).unwrap_or_default());
                (
                    COLUMN_PATTERN.is_match(&masked),
                    ROW_PATTERN.is_match(&masked),
                )
            });
            let has_column = column;
            let has_row = !column && row;

            if has_column || has_row {
                self.sheet_cells
                    .entry(sheet.sheet_index)
                    .or_default()
                    .push((at.0, at.1, has_column));
            }
        }
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext, out: &mut Vec<Violation>) {
        let Some(mut cells) = self.sheet_cells.remove(&sheet.sheet_index) else {
            return;
        };
        cells.sort_unstable();

        // Split by type
        let column_cells: Vec<(u32, u32)> =
            cells.iter().filter(|c| c.2).map(|c| (c.0, c.1)).collect();
        let row_cells: Vec<(u32, u32)> =
            cells.iter().filter(|c| !c.2).map(|c| (c.0, c.1)).collect();

        for (is_column, group) in [(true, &column_cells), (false, &row_cells)] {
            if group.is_empty() {
                continue;
            }
            let ranges = find_contiguous_ranges(group);
            for range in ranges {
                let bbox = bounding_box(&range);
                out.push(Violation::with_data(
                    RuleId::Ref307,
                    ViolationScope::Sheet(sheet.sheet_index),
                    WholeRefData {
                        is_column,
                        range: bbox,
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
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::collect_violations;
    use std::collections::HashMap;

    #[test]
    fn test_whole_column_reference() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=SUM(A:A)")),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            ..Default::default()
        };

        let mut rule = WholeColumnRowRefsRule::new();
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref307);
    }

    #[test]
    fn test_whole_row_reference() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=SUM(1:1)")),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            ..Default::default()
        };

        let mut rule = WholeColumnRowRefsRule::new();
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref307);
    }

    #[test]
    fn test_bounded_reference() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=SUM(A1:A10)")),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            ..Default::default()
        };

        let mut rule = WholeColumnRowRefsRule::new();
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert!(violations.is_empty());
    }

    #[test]
    fn whole_lines_come_from_the_parsed_formula_not_from_text() {
        assert_eq!(whole_line_refs("=SUM(A:A)"), Some((true, false)));
        assert_eq!(whole_line_refs("=SUM(Data!$B:$D)+1"), Some((true, false)));
        assert_eq!(whole_line_refs("=SUM(2:2)"), Some((false, true)));
        // Text that looks like a whole column or row is text.
        assert_eq!(
            whole_line_refs("=TEXT(A1,\"10:30\")&\"A:A\""),
            Some((false, false))
        );
        assert_eq!(whole_line_refs("=SUM(A1:A10)"), Some((false, false)));
    }
}

//! REF304: Large used range
//!
//! A sheet whose used range (the dimension the file declares for the sheet) reaches well past its
//! last real content. Stray formatting or cleared cells inflate the range, which makes the file
//! larger, slows loading, and makes "go to the last cell", printing and whole-range operations
//! land far from the data.
//!
//! # Reports
//!
//! One violation per sheet (the unit is the sheet) when the declared used range has more than 2
//! rows beyond the last row with data, or more than 2 columns beyond the last column with data.
//! The last row and the last column are taken independently (the maximum over all data cells), so
//! the "last data cell" in the message is that corner, not necessarily an existing cell. A cell
//! counts as data when its value is not empty; text that is only whitespace does not count.
//!
//! # Does not report
//!
//! - A sheet with no data cells, or without used-range metadata: nothing to compare (an empty
//!   sheet is REF303).
//! - A formula cell whose saved value is empty is not data here, even though the message says
//!   "data/formula cell"; a sheet holding only such formulas is skipped.
//! - Blank gaps inside the data: REF305 (which also sees blank rows or columns at the end of the
//!   used range).
//! - Overshoots of 2 rows or columns or fewer: a small margin is normal.
//!
//! # Parameters
//!
//! None. The tolerance of 2 rows and 2 columns is a constant.
//!
//! # Severity
//!
//! Always warning.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::HashMap;

/// Threshold for rows/columns beyond data.
const THRESHOLD_ROWS: u32 = 2;
/// Threshold for columns beyond data.
const THRESHOLD_COLS: u32 = 2;

/// Per-sheet tracked max data cell: sheet_index → (max_row, max_col).
type SheetMaxMap = HashMap<u16, (u32, u32)>;

/// Rule that detects excessively large used ranges.
///
/// Uses a constant threshold of 2 rows/columns to detect when the used range
/// extends significantly beyond actual data.
pub struct LargeUsedRangeRule {
    /// Per-sheet max data cell coordinates.
    sheet_max: SheetMaxMap,
}

impl LargeUsedRangeRule {
    /// Create a new instance.
    pub fn new() -> Self {
        Self {
            sheet_max: HashMap::new(),
        }
    }
}

impl Default for LargeUsedRangeRule {
    fn default() -> Self {
        Self::new()
    }
}

/// Incident data for REF304.
#[derive(Debug)]
pub struct LargeUsedRangeData {
    /// Used range from metadata (rows, cols).
    pub used_range: (u32, u32),
    /// Last cell with actual data (row, col).
    pub last_data: (u32, u32),
}

impl ViolationData for LargeUsedRangeData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let last_used = CellReference::new(
            self.used_range.0.saturating_sub(1),
            self.used_range.1.saturating_sub(1),
        );
        let last_data = CellReference::new(self.last_data.0, self.last_data.1);
        format!(
            "Used range extends beyond data: last used cell {last_used}, last data/formula cell {last_data} (threshold: {THRESHOLD_ROWS}/{THRESHOLD_COLS} rows/cols)",
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Check whether a cell value should count as real data.
fn is_data_cell(cell: &Cell) -> bool {
    match &cell.value {
        crate::reader::workbook::CellValue::Empty => false,
        crate::reader::workbook::CellValue::Text(s) => !s.trim().is_empty(),
        _ => true,
    }
}

impl WalkerRule for LargeUsedRangeRule {
    fn id(&self) -> RuleId {
        RuleId::Ref304
    }

    fn name(&self) -> &'static str {
        "Large Used Range"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Reference
    }

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        at: (u32, u32),
        cell: &Cell,
        _ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        if is_data_cell(cell) {
            self.sheet_max
                .entry(sheet.sheet_index)
                .and_modify(|(max_row, max_col)| {
                    *max_row = (*max_row).max(at.0);
                    *max_col = (*max_col).max(at.1);
                })
                .or_insert((at.0, at.1));
        }
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext, out: &mut Vec<Violation>) {
        let Some(&last_data) = self.sheet_max.get(&sheet.sheet_index) else {
            return; // No data cells → skip
        };

        let Some((used_rows, used_cols)) = sheet.used_range else {
            return;
        };

        let row_diff = used_rows.saturating_sub(last_data.0 + 1);
        let col_diff = used_cols.saturating_sub(last_data.1 + 1);

        if row_diff > THRESHOLD_ROWS || col_diff > THRESHOLD_COLS {
            out.push(Violation::with_data(
                RuleId::Ref304,
                ViolationScope::Sheet(sheet.sheet_index),
                LargeUsedRangeData {
                    used_range: (used_rows, used_cols),
                    last_data,
                },
                Severity::Warning,
            ));
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
    fn test_large_used_range() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                value: CellValue::Number(1.0),
            },
        );
        cells.insert(
            (1, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                value: CellValue::Number(2.0),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            used_range: Some((50, 30)),
            ..Default::default()
        };

        let mut rule = LargeUsedRangeRule::new();
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref304);
    }

    #[test]
    fn test_no_violation_within_threshold() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                value: CellValue::Number(1.0),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            used_range: Some((2, 2)), // Only 1 row/col beyond data
            ..Default::default()
        };

        let mut rule = LargeUsedRangeRule::new();
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert!(violations.is_empty());
    }

    #[test]
    fn test_empty_sheet_no_violation() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: Default::default(),
            used_range: Some((100, 100)),
            ..Default::default()
        };

        let mut rule = LargeUsedRangeRule::new();
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert!(violations.is_empty());
    }
}

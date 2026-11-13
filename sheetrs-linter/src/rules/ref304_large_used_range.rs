//! REF304: Large used range
//!
//! A sheet whose used range (the rectangle of the cells the file keeps for the sheet: with a value, a
//! formula or only a style, up to its last hidden row and column; `Sheet::kept_area`, the same in an
//! XLSX and an ODS, the `<dimension>` is not read) reaches well past its last real content. Stray formatting or cleared cells inflate the range, which makes the file
//! larger, slows loading, and makes "go to the last cell", printing and whole-range operations
//! land far from the data.
//!
//! # Reports
//!
//! One violation per sheet (the unit is the sheet) when the used range has more than 2
//! rows beyond the last row with data, or more than 2 columns beyond the last column with data.
//! The last row and the last column are taken independently (the maximum over all data cells), so
//! the "last data cell" in the message is that corner, not necessarily an existing cell. The last data cell is the corner of
//! `Sheet::used_area` (a cell with a value or a formula, the same in both formats), so no walk of the
//! cells is needed: the two corners are compared.
//!
//! # Does not report
//!
//! - A sheet with no data cells, or without used-range metadata: nothing to compare (an empty
//!   sheet is REF303).
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
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use sheetrs_io::reader::Sheet;

/// Threshold for rows/columns beyond data.
const THRESHOLD_ROWS: u32 = 2;
/// Threshold for columns beyond data.
const THRESHOLD_COLS: u32 = 2;

/// Rule that detects excessively large used ranges.
///
/// Uses a constant threshold of 2 rows/columns to detect when the used range
/// extends significantly beyond actual data.
#[derive(Debug, Default)]
pub struct LargeUsedRangeRule;

impl LargeUsedRangeRule {
    /// Create a new instance.
    #[must_use]
    pub fn new() -> Self {
        Self
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

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext, out: &mut Vec<Violation>) {
        // Two corners, no walk of the cells: the last content cell and the last cell the file keeps.
        let Some(content) = sheet.used_area else {
            return; // No data cells → skip
        };
        let last_data = (content.last_row, content.last_col);

        let Some((used_rows, used_cols)) = sheet.kept_extent() else {
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
    use crate::rules::collect_violations;
    use sheetrs_io::reader::workbook::{Cell, CellValue, Sheet};
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
            used_area: Some(sheetrs_io::reader::UsedArea::from_a1(1, 0)),
            kept_area: Some(sheetrs_io::reader::UsedArea::from_a1(49, 29)),
            ..Default::default()
        };

        let mut rule = LargeUsedRangeRule::new();
        let mut ctx = LinterContext::default();
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
            used_area: Some(sheetrs_io::reader::UsedArea::from_a1(0, 0)),
            kept_area: Some(sheetrs_io::reader::UsedArea::from_a1(1, 1)), // Only 1 row/col beyond data
            ..Default::default()
        };

        let mut rule = LargeUsedRangeRule::new();
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert!(violations.is_empty());
    }

    #[test]
    fn test_empty_sheet_no_violation() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: sheetrs_io::reader::Grid::default(),
            kept_area: Some(sheetrs_io::reader::UsedArea::from_a1(99, 99)),
            ..Default::default()
        };

        let mut rule = LargeUsedRangeRule::new();
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert!(violations.is_empty());
    }
}

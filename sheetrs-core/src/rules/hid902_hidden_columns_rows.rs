//! HID902: Hidden Columns/Rows
//!
//! Hidden rows and columns are not seen when the sheet is read or printed, but their values still
//! feed formulas. Data can be left out of a review by accident, or hidden on purpose.
//!
//! # Reports
//!
//! One violation per run of hidden columns and one per run of hidden rows on a sheet (unit: the run,
//! scoped to the sheet). Adjacent and overlapping hidden ranges are merged first, so columns B, C and
//! D hidden one by one give one violation. The message lists the run in A1 style: `Hidden columns:
//! B:D`, `Hidden rows: 3:7`; a single column or row is named alone. Columns are reported before
//! rows.
//!
//! A column or row counts as hidden when the file marks it `hidden` (XLSX), or has
//! `table:visibility` set to `collapse` or `filter` (ODS). Rows hidden by a filter are therefore
//! reported, as are collapsed outline groups.
//!
//! # Does not report
//!
//! - Hidden worksheets, which are HID901.
//! - Columns or rows narrowed to a tiny width or height without being marked hidden, and cells
//!   hidden by a number format or font colour.
//! - Whether anything is in the hidden range: an empty hidden row or column is reported too.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always `Warning`.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::Sheet;
use crate::reader::parser_utils::column_letters;
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};

/// Rule that detects hidden columns and rows
pub struct HiddenColumnsRowsRule;

/// Incident data for HID902.
#[derive(Debug)]
pub struct HiddenColumnsRowsData {
    /// Hidden column ranges: (start_col, end_col), 0-based, inclusive.
    pub columns: Vec<(u32, u32)>,
    /// Hidden row ranges: (start_row, end_row), 0-based, inclusive.
    pub rows: Vec<(u32, u32)>,
}

impl ViolationData for HiddenColumnsRowsData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        if !self.columns.is_empty() {
            let strs: Vec<String> = self.columns.iter().map(format_column_range).collect();
            format!("Hidden columns: {}", strs.join(", "))
        } else {
            let strs: Vec<String> = self.rows.iter().map(format_row_range).collect();
            format!("Hidden rows: {}", strs.join(", "))
        }
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for HiddenColumnsRowsRule {
    fn id(&self) -> RuleId {
        RuleId::Hid902
    }

    fn name(&self) -> &'static str {
        "Hidden Rows or Columns"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Hidden
    }

    fn on_sheet_start(
        &mut self,
        sheet: &Sheet,
        _ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        for range in merge_ranges(&sheet.hidden_columns) {
            out.push(Violation::with_data(
                RuleId::Hid902,
                ViolationScope::Sheet(sheet.sheet_index),
                HiddenColumnsRowsData {
                    columns: vec![range],
                    rows: Vec::new(),
                },
                Severity::Warning,
            ));
        }

        for range in merge_ranges(&sheet.hidden_rows) {
            out.push(Violation::with_data(
                RuleId::Hid902,
                ViolationScope::Sheet(sheet.sheet_index),
                HiddenColumnsRowsData {
                    columns: Vec::new(),
                    rows: vec![range],
                },
                Severity::Warning,
            ));
        }
    }
}

/// Sort inclusive ranges and merge those that overlap or are adjacent.
fn merge_ranges(ranges: &[(u32, u32)]) -> Vec<(u32, u32)> {
    let mut sorted = ranges.to_vec();
    sorted.sort_unstable();

    let mut merged: Vec<(u32, u32)> = Vec::with_capacity(sorted.len());
    for (start, end) in sorted {
        match merged.last_mut() {
            Some(last) if start <= last.1.saturating_add(1) => last.1 = last.1.max(end),
            _ => merged.push((start, end)),
        }
    }
    merged
}

/// Format column range (e.g., "A:C" or "A")
fn format_column_range(&(start, end): &(u32, u32)) -> String {
    if start == end {
        return column_letters(start).to_string();
    }
    format!("{}:{}", column_letters(start), column_letters(end))
}

/// Format row range (e.g., "1:3" or "1")
fn format_row_range(&(start, end): &(u32, u32)) -> String {
    // Convert to 1-based
    if start == end {
        return (start + 1).to_string();
    }
    format!("{}:{}", start + 1, end + 1)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::Sheet;
    use crate::rules::collect_violations;

    #[test]
    fn test_hidden_columns() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: Default::default(),
            hidden_columns: vec![(0, 1), (2, 2), (5, 5)], // A:B, C, F
            ..Default::default()
        };

        let mut rule = HiddenColumnsRowsRule;
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_start(&sheet, &mut ctx, out));

        let messages: Vec<String> = violations.iter().map(Violation::message).collect();
        assert_eq!(messages, ["Hidden columns: A:C", "Hidden columns: F"]);
        assert_eq!(violations[0].rule_id, RuleId::Hid902);
    }

    #[test]
    fn test_hidden_rows() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: Default::default(),
            hidden_rows: vec![(10, 11), (0, 2), (1, 1)], // 11:12, 1:3, 2
            ..Default::default()
        };

        let mut rule = HiddenColumnsRowsRule;
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_start(&sheet, &mut ctx, out));

        let messages: Vec<String> = violations.iter().map(Violation::message).collect();
        assert_eq!(messages, ["Hidden rows: 1:3", "Hidden rows: 11:12"]);
    }
}

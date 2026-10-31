//! REF303: Empty sheets
//!
//! An empty sheet is clutter: it adds a tab to navigate and is often the remains of a deleted
//! analysis or a sheet created by accident. Removing it makes the structure of the workbook clearer.
//!
//! # Reports
//!
//! One violation per sheet (unit: the sheet) whose cell collection is empty, that is, the reader found
//! no stored cell in it: no value, no formula and no formatted blank cell.
//!
//! # Does not report
//!
//! - Sheets with at least one stored cell of any kind. A sheet holding only a space, an empty string
//!   or blank cells that carry formatting (which the reader keeps) is not empty here. Oversized used
//!   ranges and blank rows or columns are REF304 and REF305.
//! - Hidden or unreferenced sheets that have content: HID901 and REF306.
//! - Whether another sheet refers to the sheet. The previous description said "no incoming
//!   references", but the code checks only the cells of the sheet: an empty sheet that formulas refer
//!   to is still reported.
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
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};

/// Rule that detects completely empty sheets
pub struct EmptySheetsRule;

/// Incident data for REF303.
#[derive(Debug)]
pub struct EmptySheetData {
    /// 0-based sheet index of the empty sheet.
    pub sheet_index: u16,
}

impl ViolationData for EmptySheetData {
    fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        let name = ctx
            .workbook
            .sheet_name_by_index(self.sheet_index)
            .unwrap_or("Unknown");
        format!("Sheet '{name}' is completely empty")
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for EmptySheetsRule {
    fn id(&self) -> RuleId {
        RuleId::Ref303
    }

    fn name(&self) -> &'static str {
        "Empty Sheet"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Reference
    }

    fn on_sheet_start(
        &mut self,
        sheet: &Sheet,
        _ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        if sheet.cells.is_empty() {
            out.push(Violation::with_data(
                RuleId::Ref303,
                ViolationScope::Sheet(sheet.sheet_index),
                EmptySheetData {
                    sheet_index: sheet.sheet_index,
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
    use crate::rules::LinterContext;
    use crate::rules::collect_violations;
    use std::collections::HashMap;
    use std::sync::Arc;

    #[test]
    fn test_empty_sheet_detected() {
        let empty_sheet = Sheet {
            name: "Empty".to_string(),
            sheet_index: 2,
            cells: Default::default(),
            used_range: None,
            ..Default::default()
        };

        let mut rule = EmptySheetsRule;
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_start(&empty_sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref303);
    }

    #[test]
    fn test_non_empty_sheet_not_flagged() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                value: CellValue::Text(Arc::from("Data")),
            },
        );

        let sheet = Sheet {
            name: "Main".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let mut rule = EmptySheetsRule;
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_start(&sheet, &mut ctx, out));

        assert!(violations.is_empty());
    }
}

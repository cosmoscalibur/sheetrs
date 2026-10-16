//! REF303: Empty sheets detection
//!
//! Description: Completely empty sheets with no content, formulas, or incoming references.

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
            cells: HashMap::new(),
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
                row: 0,
                col: 0,
                value: CellValue::Text(Arc::from("Data")),
            },
        );

        let sheet = Sheet {
            name: "Main".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let mut rule = EmptySheetsRule;
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_start(&sheet, &mut ctx, out));

        assert!(violations.is_empty());
    }
}

//! REF306: Unused sheets detection
//!
//! Description: Check for sheets that are not referenced by any other part of the workbook.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::Workbook;
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};

/// Rule that detects unused (standalone) sheets
pub struct UnusedSheetsRule;

/// Incident data for REF306.
#[derive(Debug)]
pub struct UnusedSheetData {
    /// 0-based sheet index of the unused sheet.
    pub sheet_index: u16,
}

impl ViolationData for UnusedSheetData {
    fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        let name = ctx
            .workbook
            .sheet_name_by_index(self.sheet_index)
            .unwrap_or("Unknown");
        format!("Sheet '{name}' is not referenced by any other sheet")
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for UnusedSheetsRule {
    fn id(&self) -> RuleId {
        RuleId::Ref306
    }

    fn name(&self) -> &'static str {
        "Unused Sheet"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Reference
    }

    fn on_workbook_end(&self, workbook: &Workbook, ctx: &mut LinterContext) -> Vec<Violation> {
        let mut violations = Vec::new();

        // Use referenced_sheets populated by calc202's on_cell
        for sheet in &workbook.sheets {
            let is_only_sheet = workbook.sheets.len() == 1;
            let is_referenced = ctx.referenced_sheets.contains(&sheet.sheet_index);

            if !is_only_sheet && !is_referenced {
                violations.push(Violation::with_data(
                    RuleId::Ref306,
                    ViolationScope::Sheet(sheet.sheet_index),
                    UnusedSheetData {
                        sheet_index: sheet.sheet_index,
                    },
                    Severity::Warning,
                ));
            }
        }

        violations
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::LinterContext;
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_unreferenced_sheet_flagged() {
        let sheet1 = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: HashMap::new(),
            ..Default::default()
        };

        let mut cells2 = HashMap::new();
        cells2.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Number(42.0),
            },
        );

        let sheet2 = Sheet {
            name: "Sheet2".to_string(),
            sheet_index: 1,
            cells: cells2,
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet1, sheet2],
            ..Default::default()
        };

        let rule = UnusedSheetsRule;
        let mut ctx = LinterContext::default();

        // Simulate: Sheet1 (index 0) is referenced by another sheet
        ctx.referenced_sheets.insert(0);

        let violations = rule.on_workbook_end(&workbook, &mut ctx);

        // Sheet2 (index 1) should be flagged as unreferenced
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref306);
    }

    #[test]
    fn test_single_sheet_not_flagged() {
        let sheet = Sheet {
            name: "OnlySheet".to_string(),
            sheet_index: 0,
            cells: HashMap::new(),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        let rule = UnusedSheetsRule;
        let mut ctx = LinterContext::default();
        let violations = rule.on_workbook_end(&workbook, &mut ctx);

        assert!(violations.is_empty());
    }
}

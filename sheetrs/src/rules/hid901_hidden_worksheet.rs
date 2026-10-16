//! HID901: Hidden Worksheet detection

use super::{LinterContext, WalkerRule};
use crate::reader::Sheet;
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};

/// Rule that detects hidden worksheets
///
/// Hidden sheets can sometimes contain sensitive data or deprecated logic that should be removed.
pub struct HiddenWorksheetRule;

/// Incident data for HID901.
#[derive(Debug)]
pub struct HiddenWorksheetData {
    /// 0-based sheet index of the hidden sheet.
    pub sheet_index: u16,
}

impl ViolationData for HiddenWorksheetData {
    fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        let name = ctx
            .workbook
            .sheet_name_by_index(self.sheet_index)
            .unwrap_or("Unknown");
        format!("Hidden sheet: {name}")
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for HiddenWorksheetRule {
    fn id(&self) -> RuleId {
        RuleId::Hid901
    }

    fn name(&self) -> &'static str {
        "Hidden Worksheet"
    }

    fn category(&self) -> super::RuleCategory {
        super::RuleCategory::Hidden
    }

    fn on_sheet_start(
        &mut self,
        sheet: &Sheet,
        _ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        if !sheet.visible {
            out.push(Violation::with_data(
                RuleId::Hid901,
                ViolationScope::Sheet(sheet.sheet_index),
                HiddenWorksheetData {
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
    use crate::reader::workbook::Sheet;
    use crate::rules::collect_violations;

    #[test]
    fn test_hidden_sheet() {
        let hidden_sheet = Sheet {
            name: "HiddenSheet1".to_string(),
            sheet_index: 1,
            visible: false,
            ..Default::default()
        };

        let mut rule = HiddenWorksheetRule;
        let mut ctx = LinterContext::default();
        let violations =
            collect_violations(|out| rule.on_sheet_start(&hidden_sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Hid901);
        assert_eq!(violations[0].scope, ViolationScope::Sheet(1));
        assert!(violations[0].message().contains("Hidden sheet"));
    }

    #[test]
    fn test_visible_sheet() {
        let visible_sheet = Sheet::new("Visible".to_string(), 0);

        let mut rule = HiddenWorksheetRule;
        let mut ctx = LinterContext::default();
        let violations =
            collect_violations(|out| rule.on_sheet_start(&visible_sheet, &mut ctx, out));

        assert_eq!(violations.len(), 0);
    }
}

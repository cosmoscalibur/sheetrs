//! HID901: Hidden Worksheet
//!
//! A hidden sheet is easy to forget and hard for a reader to find. It can hold intermediate
//! calculations nobody reviews, data that should not be shared, or leftovers from earlier versions,
//! and it still takes part in the results.
//!
//! # Reports
//!
//! One violation per hidden worksheet (unit: the sheet). The message gives the sheet name. A sheet
//! is hidden when the workbook marks it `hidden` or `veryHidden` (XLSX) or when its table style
//! is hidden (ODS). The two XLSX states are told apart in the message: a very hidden sheet does
//! not appear in the "Unhide" list of Excel, so nobody can find it from the menu. Both are
//! warnings.
//!
//! # Does not report
//!
//! - Hidden rows and columns, which are HID902.
//! - Sheets that are visible but off screen, tiny, or covered; only the sheet's visibility is
//!   read.
//! - What the sheet contains; an empty hidden sheet is reported like any other.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always `Warning`.

use super::{LinterContext, WalkerRule};
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};
use sheetrs_core::reader::{Sheet, SheetState};

/// Rule that detects hidden worksheets
///
/// Hidden sheets can sometimes contain sensitive data or deprecated logic that should be removed.
#[derive(Debug)]
pub struct HiddenWorksheetRule;

/// Incident data for HID901.
#[derive(Debug)]
pub struct HiddenWorksheetData {
    /// 0-based sheet index of the hidden sheet.
    pub sheet_index: u16,
    /// `Hidden` or `VeryHidden`.
    pub state: SheetState,
}

impl ViolationData for HiddenWorksheetData {
    fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        let name = ctx
            .workbook
            .sheet_name_by_index(self.sheet_index)
            .unwrap_or("Unknown");
        match self.state {
            SheetState::VeryHidden => format!(
                "Very hidden sheet: {name} (it does not appear in the list of sheets to unhide)"
            ),
            SheetState::Hidden | SheetState::Visible => format!("Hidden sheet: {name}"),
        }
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
        match sheet.state {
            SheetState::Visible => {}
            state @ (SheetState::Hidden | SheetState::VeryHidden) => {
                out.push(Violation::with_data(
                    RuleId::Hid901,
                    ViolationScope::Sheet(sheet.sheet_index),
                    HiddenWorksheetData {
                        sheet_index: sheet.sheet_index,
                        state,
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
    use crate::rules::collect_violations;
    use sheetrs_core::reader::workbook::{Sheet, SheetState};

    #[test]
    fn test_hidden_sheet() {
        let hidden_sheet = Sheet {
            name: "HiddenSheet1".to_string(),
            sheet_index: 1,
            state: SheetState::Hidden,
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
    fn a_very_hidden_sheet_is_told_apart_from_a_hidden_one() {
        let sheet = Sheet {
            name: "Secret".to_string(),
            sheet_index: 0,
            state: SheetState::VeryHidden,
            ..Default::default()
        };

        let mut rule = HiddenWorksheetRule;
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_start(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].severity, Severity::Warning);
        assert!(violations[0].message().starts_with("Very hidden sheet"));
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

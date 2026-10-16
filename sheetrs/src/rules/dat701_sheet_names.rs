//! DATA701: Non-descriptive sheet names

use super::{LinterContext, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::Sheet;
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};

/// Rule that detects non-descriptive sheet names (e.g. Sheet1, Sheet2)
///
/// Generic names make it hard to understand the purpose of a worksheet.
#[derive(Default)]
pub struct NonDescriptiveSheetNameRule {
    config: LinterConfig,
}

/// Incident data for DAT701.
#[derive(Debug)]
pub struct NonDescriptiveSheetNameData {
    /// 0-based sheet index.
    pub sheet_index: u16,
}

impl ViolationData for NonDescriptiveSheetNameData {
    fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        let name = ctx
            .workbook
            .sheet_name_by_index(self.sheet_index)
            .unwrap_or("Unknown");
        format!("Non-descriptive sheet name '{name}'")
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl NonDescriptiveSheetNameRule {
    /// Create a new instance with optional configuration
    pub fn new(config: &LinterConfig) -> Self {
        Self {
            config: config.clone(),
        }
    }
}

impl WalkerRule for NonDescriptiveSheetNameRule {
    fn id(&self) -> RuleId {
        RuleId::Data701
    }

    fn name(&self) -> &'static str {
        "Non-Descriptive Sheet Name"
    }

    fn category(&self) -> super::RuleCategory {
        super::RuleCategory::Data
    }

    fn on_sheet_start(
        &mut self,
        sheet: &Sheet,
        _ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        let normalized_name = sheet.name.to_lowercase();
        let patterns = self
            .config
            .get_param_array("avoid_sheet_names", Some(&sheet.name))
            .unwrap_or_else(|| vec!["sheet".to_string(), "copy".to_string()]);

        for pattern in &patterns {
            if normalized_name.contains(pattern) {
                out.push(Violation::with_data(
                    RuleId::Data701,
                    ViolationScope::Sheet(sheet.sheet_index),
                    NonDescriptiveSheetNameData {
                        sheet_index: sheet.sheet_index,
                    },
                    Severity::Warning,
                ));
                break; // Only report once per sheet
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::rules::collect_violations;

    #[test]
    fn test_non_descriptive_sheet_name() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            ..Default::default()
        };

        let mut rule = NonDescriptiveSheetNameRule::default();
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_start(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Data701);
        assert_eq!(violations[0].scope, ViolationScope::Sheet(0));
        assert!(violations[0].message().contains("sheet"));
    }

    #[test]
    fn test_copy_sheet_name() {
        let sheet = Sheet {
            name: "Copy of Data".to_string(),
            sheet_index: 1,
            ..Default::default()
        };

        let mut rule = NonDescriptiveSheetNameRule::default();
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_start(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].scope, ViolationScope::Sheet(1));
        assert!(violations[0].message().contains("Non-descriptive"));
    }

    #[test]
    fn test_descriptive_sheet_name() {
        let sheet = Sheet {
            name: "Analysis".to_string(),
            sheet_index: 2,
            ..Default::default()
        };

        let mut rule = NonDescriptiveSheetNameRule::default();
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_start(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 0);
    }
}

//! CPX503: Excessive conditional formatting detection
//!
//! Description: High counts of conditional formatting rules can degrade workbook performance.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::Sheet;
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};

/// Rule that detects excessive conditional formatting rules
pub struct ExcessiveConditionalFormattingRule {
    config: LinterConfig,
}

impl ExcessiveConditionalFormattingRule {
    /// Create a new instance with optional configuration
    pub fn new(config: &LinterConfig) -> Self {
        Self {
            config: config.clone(),
        }
    }
}

/// Incident data for CPX503.
#[derive(Debug)]
pub struct ExcessiveConditionalFormattingData {
    /// Number of conditional formatting rules found.
    pub count: usize,
    /// Affected ranges as raw strings.
    pub ranges: Vec<String>,
}

impl ViolationData for ExcessiveConditionalFormattingData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        if self.ranges.is_empty() {
            format!("Sheet has {} conditional formatting rules", self.count)
        } else {
            format!(
                "Sheet has {} conditional formatting rules. Ranges: {}",
                self.count,
                self.ranges.join(", ")
            )
        }
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for ExcessiveConditionalFormattingRule {
    fn id(&self) -> RuleId {
        RuleId::Cpx503
    }

    fn name(&self) -> &'static str {
        "Excessive Conditional Formatting"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Complexity
    }

    fn on_sheet_start(
        &mut self,
        sheet: &Sheet,
        _ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        let threshold = self
            .config
            .get_param_int("max_conditional_formatting", Some(&sheet.name))
            .unwrap_or(5) as u32;

        let cf_count = sheet.conditional_formatting_count;

        if cf_count > threshold as usize {
            let mut ranges: Vec<String> = sheet.conditional_formatting_ranges.clone();
            ranges.sort();
            ranges.dedup();

            out.push(Violation::with_data(
                RuleId::Cpx503,
                ViolationScope::Sheet(sheet.sheet_index),
                ExcessiveConditionalFormattingData {
                    count: cf_count,
                    ranges,
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
    fn test_no_violations() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            ..Default::default()
        };

        let mut rule = ExcessiveConditionalFormattingRule::new(&LinterConfig::default());
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_start(&sheet, &mut ctx, out));

        assert!(violations.is_empty());
    }

    #[test]
    fn test_excessive_cf() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            conditional_formatting_count: 10,
            conditional_formatting_ranges: vec![
                "A1:A10".to_string(),
                "B1:B10".to_string(),
                "A1:A10".to_string(), // Duplicate
            ],
            ..Default::default()
        };

        let mut rule = ExcessiveConditionalFormattingRule::new(&LinterConfig::default());
        let mut ctx = LinterContext::default();
        let violations = collect_violations(|out| rule.on_sheet_start(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Cpx503);
    }
}

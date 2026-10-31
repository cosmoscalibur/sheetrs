//! DATA701: Non-descriptive sheet names
//!
//! A sheet called `Sheet1` or `Copy of Data` says nothing about what it holds. Default names left in
//! place and names produced by duplicating a sheet make a workbook hard to navigate and make
//! formulas such as `=Sheet3!B2` hard to read, so the reader has to open the sheet to learn its
//! purpose.
//!
//! # Reports
//!
//! One violation per sheet (unit: the sheet) whose normalized name (lowercase) contains at least one of the
//! normalized patterns as a substring. The sheet is reported once however many patterns match.
//! With the default patterns `Sheet1`, `sheet 2`, `Copy of Data` and `Data copy` are reported, and so
//! is any name that merely contains the text, such as `Worksheet` or `Photocopy`.
//!
//! # Does not report
//!
//! - Sheets whose name contains none of the patterns: the rule judges the name only, never the
//!   content, so a sheet named `Analysis` is accepted even if it is empty (see REF303 for empty
//!   sheets).
//! - Two sheets whose names differ only in case or punctuation: that is REF302.
//! - Hidden or unused sheets: those are HID901 and REF306.
//!
//! # Parameters
//!
//! - `avoid_sheet_names` (array of strings, default `["sheet", "copy"]`): the substrings to look for.
//!   Patterns and names are normalized the same way before they are compared (surrounding spaces
//!   removed, lowercase), so `Sheet` and `sheet` are the same pattern; an empty pattern is ignored.
//!   An empty array disables the check. It can be set per
//!   sheet; the sheet's own name selects the override.
//!
//! # Severity
//!
//! Always `Warning`.

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
            // The pattern is normalized as the name is: without surrounding spaces, in lowercase.
            // An empty one would match every name.
            let pattern = pattern.trim().to_lowercase();
            if !pattern.is_empty() && normalized_name.contains(&pattern) {
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
    fn patterns_are_normalized_like_names() {
        let config = LinterConfig::from_toml(
            "[global]\navoid_sheet_names = [\" Draft \", \"\", \"TEMP\"]\n",
        )
        .unwrap();
        let reported = |name: &str| {
            let sheet = Sheet {
                name: name.to_string(),
                ..Default::default()
            };
            let mut rule = NonDescriptiveSheetNameRule::new(&config);
            let mut ctx = LinterContext::default();
            collect_violations(|out| rule.on_sheet_start(&sheet, &mut ctx, out)).len()
        };
        // Capitals and surrounding spaces of a pattern do not matter; an empty pattern matches nothing.
        assert_eq!(reported("My draft v2"), 1);
        assert_eq!(reported("Temp"), 1);
        assert_eq!(reported("Analysis"), 0);
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

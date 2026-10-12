//! FILE1003: 1904 Date System detection
//!
//! Detects spreadsheets using the 1904 date system (legacy Mac compatibility).
//! This can cause date calculation issues when sharing files between systems.

use super::{LinterContext, WalkerRule};
use crate::reader::Workbook;
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};

/// Rule that identifies spreadsheets using the 1904 date system.
///
/// The 1904 date system:
/// - Was originally used by Excel on Macintosh for backward compatibility
/// - Uses January 1, 1904 as the epoch instead of January 1, 1900
/// - Can cause date values to be off by 4 years and 1 day when shared
/// - Most modern spreadsheets use the 1900 date system (default)
pub struct DateSystem1904Rule;

/// Incident data for FILE1003 — no variable fields.
#[derive(Debug)]
pub struct DateSystem1904Data;

impl ViolationData for DateSystem1904Data {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        "Workbook uses 1904 date system. This can cause date calculation \
         issues when sharing files between systems."
            .to_string()
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl DateSystem1904Rule {
    /// Creates a new rule instance.
    #[must_use]
    pub fn new() -> Self {
        Self
    }

    /// Check for 1904 date system and return violation if found.
    fn check_date_system(&self, workbook: &Workbook) -> Vec<Violation> {
        let mut violations = Vec::new();

        if workbook.date1904 {
            violations.push(Violation::with_data(
                RuleId::File1003,
                ViolationScope::Book,
                DateSystem1904Data,
                Severity::Warning,
            ));
        }

        violations
    }
}

impl Default for DateSystem1904Rule {
    fn default() -> Self {
        Self::new()
    }
}

impl WalkerRule for DateSystem1904Rule {
    fn id(&self) -> RuleId {
        RuleId::File1003
    }

    fn name(&self) -> &'static str {
        "1904 Date System"
    }

    fn category(&self) -> super::RuleCategory {
        super::RuleCategory::File
    }

    fn on_workbook_start(&self, workbook: &Workbook, _ctx: &mut LinterContext) -> Vec<Violation> {
        self.check_date_system(workbook)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_date_system_no_violation() {
        let workbook = Workbook {
            date1904: false,
            ..Default::default()
        };

        let rule = DateSystem1904Rule::new();
        let violations = rule.check_date_system(&workbook);
        assert!(violations.is_empty());
    }

    #[test]
    fn test_1904_date_system_violation() {
        let workbook = Workbook {
            date1904: true,
            ..Default::default()
        };

        let rule = DateSystem1904Rule::new();
        let violations = rule.check_date_system(&workbook);
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::File1003);
        assert!(violations[0].message().contains("1904 date system"));
    }

    #[test]
    fn test_default_workbook_no_violation() {
        let workbook = Workbook::default();
        let rule = DateSystem1904Rule::new();
        let violations = rule.check_date_system(&workbook);
        assert!(violations.is_empty());
    }
}

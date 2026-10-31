//! FILE1003: 1904 Date System
//!
//! In the 1904 date system, day 0 is 1 January 1904 instead of 1900. The same stored number is
//! a date four years and one day later, so dates copied or linked between a 1904 and a 1900
//! workbook are silently shifted, and date arithmetic against other files goes wrong.
//!
//! # Reports
//!
//! One violation for the book when the workbook is set to the 1904 date system: `date1904` is
//! `1` or `true` on the workbook properties of an XLSX file, or the ODS `table:null-date` is a
//! date in 1904. The message is the same for every file.
//!
//! # Does not report
//!
//! - Other non-default null dates in ODS (the value must start with `1904`), and any date system
//!   that is not stated in the file.
//! - Dates that look wrong in cells; the rule does not read cell values, only the setting.
//! - A 1900 date system, the default.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always `Warning`.

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

    fn on_workbook_start(
        &mut self,
        workbook: &Workbook,
        _ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        if workbook.date1904 {
            out.push(Violation::with_data(
                RuleId::File1003,
                ViolationScope::Book,
                DateSystem1904Data,
                Severity::Warning,
            ));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::rules::collect_violations;

    #[test]
    fn test_default_date_system_no_violation() {
        let workbook = Workbook {
            date1904: false,
            ..Default::default()
        };

        let mut rule = DateSystem1904Rule::new();
        let violations = collect_violations(|out| {
            rule.on_workbook_start(&workbook, &mut LinterContext::default(), out)
        });
        assert!(violations.is_empty());
    }

    #[test]
    fn test_1904_date_system_violation() {
        let workbook = Workbook {
            date1904: true,
            ..Default::default()
        };

        let mut rule = DateSystem1904Rule::new();
        let violations = collect_violations(|out| {
            rule.on_workbook_start(&workbook, &mut LinterContext::default(), out)
        });
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::File1003);
        assert!(violations[0].message().contains("1904 date system"));
    }

    #[test]
    fn test_default_workbook_no_violation() {
        let workbook = Workbook::default();
        let mut rule = DateSystem1904Rule::new();
        let violations = collect_violations(|out| {
            rule.on_workbook_start(&workbook, &mut LinterContext::default(), out)
        });
        assert!(violations.is_empty());
    }
}

//! FILE1002: Old Spreadsheet detection
//!
//! Detects spreadsheets that haven't been modified in a long time,
//! which may indicate stale or outdated data requiring review.

use super::{LinterContext, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};
use chrono::Utc;

/// Default maximum age threshold in days
const DEFAULT_MAX_AGE_DAYS: i64 = 365;

/// Rule that identifies spreadsheets that are older than a threshold.
///
/// Old spreadsheets may contain:
/// - Outdated data requiring review or update
/// - Stale references or links
/// - Information no longer relevant to current processes
pub struct OldSpreadsheetRule {
    /// Maximum age in days
    max_age_days: i64,
}

/// Incident data for FILE1002.
#[derive(Debug)]
pub struct OldSpreadsheetData {
    /// Observed age in days.
    pub age_days: i64,
}

impl ViolationData for OldSpreadsheetData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!("File is {} days old", self.age_days)
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl OldSpreadsheetRule {
    /// Creates a new rule from config, reading `max_age_days` parameter.
    #[must_use]
    pub fn new(config: &LinterConfig) -> Self {
        let max_days = config
            .get_param_int("max_age_days", None)
            .unwrap_or(DEFAULT_MAX_AGE_DAYS);
        Self {
            max_age_days: max_days,
        }
    }

    /// Check file age and return violation if exceeded.
    fn check_age(&self, workbook: &Workbook) -> Vec<Violation> {
        let mut violations = Vec::new();

        // modified_date is already parsed by parser (Parse, don't validate)
        if let Some(modified) = workbook.modified_date {
            let age = Utc::now().signed_duration_since(modified);
            let age_days = age.num_days();

            if age_days > self.max_age_days {
                violations.push(Violation::with_data(
                    RuleId::File1002,
                    ViolationScope::Book,
                    OldSpreadsheetData { age_days },
                    Severity::Info,
                ));
            }
        }

        violations
    }
}

impl Default for OldSpreadsheetRule {
    fn default() -> Self {
        Self {
            max_age_days: DEFAULT_MAX_AGE_DAYS,
        }
    }
}

impl WalkerRule for OldSpreadsheetRule {
    fn id(&self) -> RuleId {
        RuleId::File1002
    }

    fn name(&self) -> &'static str {
        "Old Spreadsheet Format"
    }

    fn category(&self) -> super::RuleCategory {
        super::RuleCategory::File
    }

    fn on_workbook_start(&self, workbook: &Workbook, _ctx: &mut LinterContext) -> Vec<Violation> {
        self.check_age(workbook)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::{DateTime, Utc};

    #[test]
    fn test_recent_file_no_violation() {
        let workbook = Workbook::default();
        let rule = OldSpreadsheetRule::default();
        let violations = rule.check_age(&workbook);
        assert!(violations.is_empty());
    }

    #[test]
    fn test_old_file_violation() {
        // Create a date more than 365 days ago
        let old_date: DateTime<Utc> = "2020-01-01T00:00:00Z".parse().unwrap();
        let workbook = Workbook {
            modified_date: Some(old_date),
            ..Default::default()
        };

        let rule = OldSpreadsheetRule::default();
        let violations = rule.check_age(&workbook);
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::File1002);
        assert!(violations[0].message().contains("days old"));
    }

    #[test]
    fn test_missing_date_no_violation() {
        let workbook = Workbook::default();
        let rule = OldSpreadsheetRule::default();
        let violations = rule.check_age(&workbook);
        assert!(violations.is_empty());
    }
}

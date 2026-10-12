//! FILE1001: Large File Size detection
//!
//! Detects physical size bloat indicating instability or excessive data.
//! Uses file_size_bytes metadata from Workbook (Parse-Dont-Validate pattern).

use super::{LinterContext, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};

/// Default maximum file size threshold in megabytes
const DEFAULT_MAX_FILE_SIZE_MB: i64 = 10;

/// Rule that identifies excessively large physical file sizes.
///
/// Large spreadsheet files may indicate:
/// - Excessive data storage beyond intended use
/// - Embedded objects or images bloating file size
/// - Potential instability or corruption risk
/// - Performance issues when opening/saving
pub struct LargeFileSizeRule {
    /// Maximum file size in bytes
    max_file_size_bytes: u64,
}

/// Incident data for FILE1001.
#[derive(Debug)]
pub struct LargeFileSizeData {
    /// Observed file size in bytes.
    pub size_bytes: u64,
}

impl ViolationData for LargeFileSizeData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let size_mb = self.size_bytes as f64 / (1024.0 * 1024.0);
        format!("File size {size_mb:.2} MB exceeds limit")
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl LargeFileSizeRule {
    /// Creates a new rule from config, reading `max_file_size_mb` parameter.
    #[must_use]
    pub fn new(config: &LinterConfig) -> Self {
        let max_mb = config
            .get_param_int("max_file_size_mb", None)
            .unwrap_or(DEFAULT_MAX_FILE_SIZE_MB);
        Self {
            max_file_size_bytes: (max_mb * 1024 * 1024) as u64,
        }
    }

    /// Check file size and return violation if exceeded.
    /// Uses file_size_bytes from Workbook metadata (pre-extracted by parser).
    fn check_size(&self, workbook: &Workbook) -> Vec<Violation> {
        let mut violations = Vec::new();

        let file_size = workbook.file_size_bytes;
        if file_size > self.max_file_size_bytes {
            violations.push(Violation::with_data(
                RuleId::File1001,
                ViolationScope::Book,
                LargeFileSizeData {
                    size_bytes: file_size,
                },
                Severity::Warning,
            ));
        }

        violations
    }
}

impl Default for LargeFileSizeRule {
    fn default() -> Self {
        Self {
            max_file_size_bytes: (DEFAULT_MAX_FILE_SIZE_MB * 1024 * 1024) as u64,
        }
    }
}

impl WalkerRule for LargeFileSizeRule {
    fn id(&self) -> RuleId {
        RuleId::File1001
    }

    fn name(&self) -> &'static str {
        "Large File Size"
    }

    fn category(&self) -> super::RuleCategory {
        super::RuleCategory::File
    }

    fn on_workbook_start(
        &mut self,
        workbook: &Workbook,
        _ctx: &mut LinterContext,
    ) -> Vec<Violation> {
        self.check_size(workbook)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_small_file_no_violation() {
        let workbook = Workbook {
            file_size_bytes: 1024, // 1 KB
            ..Default::default()
        };

        let rule = LargeFileSizeRule::default();
        let violations = rule.check_size(&workbook);
        assert!(violations.is_empty());
    }

    #[test]
    fn test_large_file_violation() {
        let rule = LargeFileSizeRule {
            max_file_size_bytes: 1024, // 1 KB threshold
        };

        let workbook = Workbook {
            file_size_bytes: 2048, // 2 KB
            ..Default::default()
        };

        let violations = rule.check_size(&workbook);
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::File1001);
        assert!(violations[0].message().contains("exceeds limit"));
    }

    #[test]
    fn test_zero_size_no_violation() {
        let workbook = Workbook::default();
        let rule = LargeFileSizeRule::default();
        let violations = rule.check_size(&workbook);
        assert!(violations.is_empty());
    }
}

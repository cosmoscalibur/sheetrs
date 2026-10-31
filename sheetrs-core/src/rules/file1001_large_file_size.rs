//! FILE1001: Large File Size
//!
//! An unusually large workbook is slow to open, save, send and diff, and is more likely to be
//! corrupted. Size often comes from forgotten data, embedded images or objects, or unused
//! formatting rather than from the model itself.
//!
//! # Reports
//!
//! One violation for the book when the size of the file on disk exceeds the limit. The size is
//! the stored (compressed) size in bytes, not the size of the data inside; the message shows it
//! in MB (1 MB is 1024 x 1024 bytes) and does not repeat the limit. The comparison is strictly
//! greater than, so a file exactly at the limit is not reported.
//!
//! # Does not report
//!
//! - What makes the file large. The rule only states that it is; sheet and formatting size rules
//!   (for example REF304, CPX503) point at causes.
//! - A file whose size is not known: the size is read from the file system and counts as 0 when it
//!   cannot be read, so it is never reported.
//!
//! # Parameters
//!
//! - `max_file_size_mb` (integer): limit in MB. Default `10`.
//!
//! # Severity
//!
//! Always `Warning`.

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
        out: &mut Vec<Violation>,
    ) {
        let file_size = workbook.file_size_bytes;
        if file_size > self.max_file_size_bytes {
            out.push(Violation::with_data(
                RuleId::File1001,
                ViolationScope::Book,
                LargeFileSizeData {
                    size_bytes: file_size,
                },
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
    fn test_small_file_no_violation() {
        let workbook = Workbook {
            file_size_bytes: 1024, // 1 KB
            ..Default::default()
        };

        let mut rule = LargeFileSizeRule::default();
        let violations = collect_violations(|out| {
            rule.on_workbook_start(&workbook, &mut LinterContext::default(), out)
        });
        assert!(violations.is_empty());
    }

    #[test]
    fn test_large_file_violation() {
        let mut rule = LargeFileSizeRule {
            max_file_size_bytes: 1024, // 1 KB threshold
        };

        let workbook = Workbook {
            file_size_bytes: 2048, // 2 KB
            ..Default::default()
        };

        let violations = collect_violations(|out| {
            rule.on_workbook_start(&workbook, &mut LinterContext::default(), out)
        });
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::File1001);
        assert!(violations[0].message().contains("exceeds limit"));
    }

    #[test]
    fn test_zero_size_no_violation() {
        let workbook = Workbook::default();
        let mut rule = LargeFileSizeRule::default();
        let violations = collect_violations(|out| {
            rule.on_workbook_start(&workbook, &mut LinterContext::default(), out)
        });
        assert!(violations.is_empty());
    }
}

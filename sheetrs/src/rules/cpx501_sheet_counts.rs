//! CPX501: Excessive sheet counts
//!
//! Description: Checks if the workbook has an excessive number of sheets, which complicates navigation.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};

/// Rule that checks if the workbook has an excessive number of sheets.
pub struct ExcessiveSheetCountsRule {
    threshold: u32,
}

/// Incident data for CPX501.
#[derive(Debug)]
pub struct ExcessiveSheetCountsData {
    /// Observed sheet count.
    pub count: u32,
}

impl ViolationData for ExcessiveSheetCountsData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!("Workbook has {} sheets", self.count)
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl ExcessiveSheetCountsRule {
    pub fn new(config: &LinterConfig) -> Self {
        let threshold = config.get_param_int("max_sheets", None).unwrap_or(50);

        Self {
            threshold: threshold as u32,
        }
    }
}

impl Default for ExcessiveSheetCountsRule {
    fn default() -> Self {
        Self { threshold: 50 }
    }
}

impl WalkerRule for ExcessiveSheetCountsRule {
    fn id(&self) -> RuleId {
        RuleId::Cpx501
    }

    fn name(&self) -> &'static str {
        "Sheet Counts"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Complexity
    }

    fn on_workbook_start(&self, workbook: &Workbook, _ctx: &mut LinterContext) -> Vec<Violation> {
        let mut violations = Vec::new();
        let sheet_count = workbook.sheets.len() as u32;

        if sheet_count > self.threshold {
            violations.push(Violation::with_data(
                RuleId::Cpx501,
                ViolationScope::Book,
                ExcessiveSheetCountsData { count: sheet_count },
                Severity::Warning,
            ));
        }

        violations
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::Sheet;
    use crate::rules::walker::WorkbookWalker;
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_excessive_sheet_counts() {
        let mut sheets = Vec::new();
        for i in 0..60 {
            sheets.push(Sheet {
                name: format!("Sheet{i}"),
                sheet_index: i as u16,
                cells: HashMap::new(),
                used_range: None,
                hidden_columns: Vec::new(),
                hidden_rows: Vec::new(),
                merged_cells: Vec::new(),
                sheet_path: None,
                formula_parsing_error: None,
                conditional_formatting_count: 0,
                conditional_formatting_ranges: Vec::new(),
                visible: true,
            });
        }

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets,
            ..Default::default()
        };

        let rule = ExcessiveSheetCountsRule::default();
        let rules: Vec<Box<dyn WalkerRule>> = vec![Box::new(rule)];
        let walker = WorkbookWalker::new(&workbook, rules);
        let violations = walker.walk();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Cpx501);
        assert!(violations[0].message().contains("60 sheets"));
    }
}

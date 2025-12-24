//! SM001: Excessive sheet counts

use super::{LinterRule, RuleCategory};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;

pub struct ExcessiveSheetCountsRule {
    threshold: u32,
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

impl LinterRule for ExcessiveSheetCountsRule {
    fn id(&self) -> &str {
        "SM001"
    }

    fn name(&self) -> &str {
        "Excessive sheet counts"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::StructuralAndMaintainability
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();
        let sheet_count = workbook.sheets.len() as u32;

        if sheet_count > self.threshold {
            violations.push(Violation::new(
                self.id(),
                ViolationScope::Book,
                format!(
                    "Workbook has {} sheets (threshold: {})",
                    sheet_count, self.threshold
                ),
                Severity::Warning,
            ));
        }

        Ok(violations)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::Sheet;
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_excessive_sheet_counts() {
        let mut sheets = Vec::new();
        for i in 0..60 {
            sheets.push(Sheet {
                name: format!("Sheet{}", i),
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
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
            external_workbooks: Vec::new(),
        };

        let rule = ExcessiveSheetCountsRule::default();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "SM001");
        assert!(violations[0].message.contains("60 sheets"));
    }
}

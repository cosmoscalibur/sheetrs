//! SEC002: Hidden sheets detection

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;

pub struct HiddenSheetsRule;

impl LinterRule for HiddenSheetsRule {
    fn id(&self) -> &str {
        "SEC002"
    }

    fn name(&self) -> &str {
        "Hidden sheets"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::SecurityAndPrivacy
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for hidden_sheet_name in &workbook.hidden_sheets {
            violations.push(Violation::new(
                self.id(),
                ViolationScope::Book,
                format!("Hidden sheet: {}", hidden_sheet_name),
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
    fn test_hidden_sheets() {
        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![Sheet {
                name: "Visible".to_string(),
                cells: HashMap::new(),
                used_range: None,
                hidden_columns: Vec::new(),
                hidden_rows: Vec::new(),
                merged_cells: Vec::new(),
                sheet_path: None,
                formula_parsing_error: None,
            }],
            defined_names: HashMap::new(),
            hidden_sheets: vec!["HiddenSheet1".to_string(), "HiddenSheet2".to_string()],
            has_macros: false,
            external_links: Vec::new(),
        };

        let rule = HiddenSheetsRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 2);
        assert_eq!(violations[0].rule_id, "SEC002");
        assert!(violations[0].message.contains("HiddenSheet1"));
        assert!(violations[1].message.contains("HiddenSheet2"));
    }

    #[test]
    fn test_no_hidden_sheets() {
        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
        };

        let rule = HiddenSheetsRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 0);
    }
}

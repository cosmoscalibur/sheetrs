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

        for sheet in &workbook.sheets {
            if !sheet.visible {
                violations.push(Violation::new(
                    self.id(),
                    ViolationScope::Book,
                    format!("Hidden sheet: {}", sheet.name),
                    Severity::Warning,
                ));
            }
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
        let visible_sheet = Sheet::new("Visible".to_string());
        let hidden_sheet1 = Sheet {
            name: "HiddenSheet1".to_string(),
            visible: false,
            ..Default::default()
        };
        let hidden_sheet2 = Sheet {
            name: "HiddenSheet2".to_string(),
            visible: false,
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![visible_sheet, hidden_sheet1, hidden_sheet2],
            ..Default::default()
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
            ..Default::default()
        };

        let rule = HiddenSheetsRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 0);
    }
}

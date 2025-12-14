//! SEC004: Macros and VBA code detection

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;

pub struct MacrosVbaRule;

impl LinterRule for MacrosVbaRule {
    fn id(&self) -> &str {
        "SEC004"
    }

    fn name(&self) -> &str {
        "Macros and VBA code"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::SecurityAndPrivacy
    }

    fn default_active(&self) -> bool {
        false
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        if workbook.has_macros {
            violations.push(Violation::new(
                self.id(),
                ViolationScope::Book,
                "Workbook contains macros or VBA code. Review for security concerns.".to_string(),
                Severity::Warning,
            ));
        }

        Ok(violations)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_workbook_with_macros() {
        let workbook = Workbook {
            path: PathBuf::from("test.xlsm"),
            sheets: vec![],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: true,
        };

        let rule = MacrosVbaRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "SEC004");
        assert!(violations[0].message.contains("macros"));
    }

    #[test]
    fn test_workbook_without_macros() {
        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
        };

        let rule = MacrosVbaRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 0);
    }
}

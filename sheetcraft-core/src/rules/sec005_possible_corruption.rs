use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::rules::{LinterRule, RuleCategory};
use crate::violation::{Severity, Violation, ViolationScope};

/// SEC005: Possible corruption file in formula parser
///
/// Detects if there were any errors during formula parsing for a sheet.
/// This indicates potential file corruption or features not supported by the parser.
pub struct PossibleCorruptionRule;

impl PossibleCorruptionRule {
    pub fn new(_config: &LinterConfig) -> Self {
        Self
    }
}

impl LinterRule for PossibleCorruptionRule {
    fn id(&self) -> &'static str {
        "SEC005"
    }

    fn name(&self) -> &'static str {
        "PossibleCorruptionRule"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::SecurityAndPrivacy
    }

    fn default_active(&self) -> bool {
        true
    }

    fn check(&self, workbook: &Workbook) -> anyhow::Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            if let Some(err_msg) = &sheet.formula_parsing_error {
                violations.push(Violation::new(
                    self.id(),
                    ViolationScope::Sheet(sheet.name.clone()),
                    format!("Possible corruption in formula parser. Error: {}", err_msg),
                    Severity::Error, // Assuming corruption is high severity
                ));
            }
        }

        Ok(violations)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::Sheet;
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_possible_corruption_detection() {
        let cells = HashMap::new();

        let sheet_normal = Sheet {
            name: "Normal".to_string(),
            cells: cells.clone(),
            used_range: None,
            hidden_columns: vec![],
            hidden_rows: vec![],
            merged_cells: vec![],
            formula_parsing_error: None,
        };

        let sheet_corrupt = Sheet {
            name: "Corrupt".to_string(),
            cells: cells.clone(),
            used_range: None,
            hidden_columns: vec![],
            hidden_rows: vec![],
            merged_cells: vec![],
            formula_parsing_error: Some("Unexpected EOF in formula expression".to_string()),
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet_normal, sheet_corrupt],
            defined_names: HashMap::new(),
            hidden_sheets: vec![],
            has_macros: false,
        };

        let config = LinterConfig::default();
        let rule = PossibleCorruptionRule::new(&config);

        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        let v = &violations[0];
        assert_eq!(v.scope, ViolationScope::Sheet("Corrupt".to_string()));
        assert!(v.message.contains("Possible corruption in formula parser"));
        assert!(v.message.contains("Unexpected EOF"));
    }
}

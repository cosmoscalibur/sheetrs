//! SM002: Duplicate sheet names

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;
use std::collections::HashMap;

pub struct DuplicateSheetNamesRule;

/// Normalize sheet name for comparison by converting to lowercase and removing non-alphanumeric characters
fn normalize_sheet_name(name: &str) -> String {
    name.to_lowercase()
        .chars()
        .filter(|c| c.is_alphanumeric())
        .collect()
}

impl LinterRule for DuplicateSheetNamesRule {
    fn id(&self) -> &str {
        "SM002"
    }

    fn name(&self) -> &str {
        "Duplicate sheet names"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::StructuralAndMaintainability
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();
        let mut name_map: HashMap<String, Vec<String>> = HashMap::new();

        // Group sheet names by their normalized version
        for sheet in &workbook.sheets {
            let normalized = normalize_sheet_name(&sheet.name);
            name_map
                .entry(normalized)
                .or_default()
                .push(sheet.name.clone());
        }

        // Find duplicates
        for (_normalized, variants) in name_map {
            if variants.len() > 1 {
                violations.push(Violation::new(
                    self.id(),
                    ViolationScope::Book,
                    format!("Confusingly similar sheet names: {}", variants.join(", ")),
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
    fn test_duplicate_sheet_names() {
        let sheets = vec![
            Sheet {
                name: "Data".to_string(),
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "data".to_string(),
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "Summary".to_string(),
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
        ];

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets,
            ..Default::default()
        };

        let rule = DuplicateSheetNamesRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "SM002");
        assert!(violations[0].message.contains("Data"));
        assert!(violations[0].message.contains("data"));
    }

    #[test]
    fn test_confusingly_similar_sheet_names() {
        let sheets = vec![
            Sheet {
                name: "Sheet1".to_string(),
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "sheet 1".to_string(),
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "Sheet-1".to_string(),
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "Data_2024".to_string(),
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
            Sheet {
                name: "data2024".to_string(),
                cells: HashMap::new(),
                used_range: None,
                ..Default::default()
            },
        ];

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets,
            ..Default::default()
        };

        let rule = DuplicateSheetNamesRule;
        let violations = rule.check(&workbook).unwrap();

        // Should detect 2 groups of duplicates:
        // 1. "Sheet1", "sheet 1", "Sheet-1" (all normalize to "sheet1")
        // 2. "Data_2024", "data2024" (both normalize to "data2024")
        assert_eq!(violations.len(), 2);

        // Check first violation contains all Sheet1 variants
        let sheet1_violation = violations.iter().find(|v| v.message.contains("Sheet1"));
        assert!(sheet1_violation.is_some());
        let msg = &sheet1_violation.unwrap().message;
        assert!(msg.contains("Sheet1"));
        assert!(msg.contains("sheet 1"));
        assert!(msg.contains("Sheet-1"));

        // Check second violation contains Data_2024 variants
        let data_violation = violations.iter().find(|v| v.message.contains("Data_2024"));
        assert!(data_violation.is_some());
        let msg = &data_violation.unwrap().message;
        assert!(msg.contains("Data_2024"));
        assert!(msg.contains("data2024"));
    }
}

//! SM002: Duplicate sheet names

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;
use std::collections::HashMap;

pub struct DuplicateSheetNamesRule;

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

    fn default_active(&self) -> bool {
        false
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();
        let mut name_map: HashMap<String, Vec<String>> = HashMap::new();

        // Group sheet names by their lowercase version
        for sheet in &workbook.sheets {
            let normalized = sheet.name.to_lowercase();
            name_map
                .entry(normalized)
                .or_insert_with(Vec::new)
                .push(sheet.name.clone());
        }

        // Find duplicates
        for (_normalized, variants) in name_map {
            if variants.len() > 1 {
                violations.push(Violation::new(
                    self.id(),
                    ViolationScope::Book,
                    format!(
                        "Duplicate sheet names (case-insensitive): {}",
                        variants.join(", ")
                    ),
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
                hidden_columns: Vec::new(),
                hidden_rows: Vec::new(),
                merged_cells: Vec::new(), formula_parsing_error: None,
            },
            Sheet {
                name: "data".to_string(),
                cells: HashMap::new(),
                used_range: None,
                hidden_columns: Vec::new(),
                hidden_rows: Vec::new(),
                merged_cells: Vec::new(), formula_parsing_error: None,
            },
            Sheet {
                name: "Summary".to_string(),
                cells: HashMap::new(),
                used_range: None,
                hidden_columns: Vec::new(),
                hidden_rows: Vec::new(),
                merged_cells: Vec::new(), formula_parsing_error: None,
            },
        ];

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets,
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
        };

        let rule = DuplicateSheetNamesRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "SM002");
        assert!(violations[0].message.contains("Data"));
        assert!(violations[0].message.contains("data"));
    }
}

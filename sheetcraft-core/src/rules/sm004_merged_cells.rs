//! SM004: Merged cells detection

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{CellReference, Severity, Violation, ViolationScope};
use anyhow::Result;

pub struct MergedCellsRule;

impl LinterRule for MergedCellsRule {
    fn id(&self) -> &str {
        "SM004"
    }

    fn name(&self) -> &str {
        "Merged cells"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::StructuralAndMaintainability
    }

    fn default_active(&self) -> bool {
        false
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            for &(start_row, start_col, end_row, end_col) in &sheet.merged_cells {
                let range_str = format_merged_range(start_row, start_col, end_row, end_col);
                violations.push(Violation::new(
                    self.id(),
                    ViolationScope::Sheet(sheet.name.clone()),
                    format!("Merged cells in range: {}", range_str),
                    Severity::Warning,
                ));
            }
        }

        Ok(violations)
    }
}

/// Format a merged cell range
fn format_merged_range(start_row: u32, start_col: u32, end_row: u32, end_col: u32) -> String {
    let start = CellReference::new(start_row, start_col);
    let end = CellReference::new(end_row, end_col);
    format!("{}:{}", start, end)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::Sheet;
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_merged_cells() {
        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![Sheet {
                name: "Sheet1".to_string(),
                cells: HashMap::new(),
                used_range: None,
                hidden_columns: Vec::new(),
                hidden_rows: Vec::new(),
                merged_cells: vec![
                    (0, 0, 0, 2), // A1:C1
                    (2, 0, 4, 0), // A3:A5
                ],
                formula_parsing_error: None,
                sheet_path: None,
            }],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
        };

        let rule = MergedCellsRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 2);
        assert_eq!(violations[0].rule_id, "SM004");
        assert!(violations[0].message.contains("A1:C1"));
        assert!(violations[1].message.contains("A3:A5"));
    }

    #[test]
    fn test_no_merged_cells() {
        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![Sheet {
                name: "Sheet1".to_string(),
                cells: HashMap::new(),
                used_range: None,
                hidden_columns: Vec::new(),
                hidden_rows: Vec::new(),
                merged_cells: Vec::new(),
                formula_parsing_error: None,
                sheet_path: None,
            }],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
        };

        let rule = MergedCellsRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 0);
    }
}

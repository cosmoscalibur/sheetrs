//! PERF001: Unused named ranges detection

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;
use std::collections::HashSet;

pub struct UnusedNamedRangesRule;

impl LinterRule for UnusedNamedRangesRule {
    fn id(&self) -> &str {
        "PERF001"
    }

    fn name(&self) -> &str {
        "Unused named ranges"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Performance
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        // Collect all named ranges
        let named_ranges: HashSet<&str> = workbook
            .defined_names
            .keys()
            .filter(|name| !name.starts_with("_xlnm."))
            .map(|s| s.as_str())
            .collect();

        // Collect all named ranges used in formulas
        let mut used_names = HashSet::new();
        for sheet in &workbook.sheets {
            for cell in sheet.all_cells() {
                if let Some(formula) = cell.value.as_formula() {
                    for name in &named_ranges {
                        if formula.contains(name) {
                            used_names.insert(*name);
                        }
                    }
                }
            }
        }

        // Report unused named ranges
        for name in named_ranges {
            if !used_names.contains(name) {
                violations.push(Violation::new(
                    self.id(),
                    ViolationScope::Book,
                    format!("Named range '{}' is defined but never used", name),
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
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_unused_named_ranges() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=UsedRange".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 1)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            sheet_path: None,
            formula_parsing_error: None,
            conditional_formatting_count: 0,
            conditional_formatting_ranges: Vec::new(),
        };

        let mut defined_names = HashMap::new();
        defined_names.insert("UsedRange".to_string(), "Sheet1!A1:B2".to_string());
        defined_names.insert("UnusedRange".to_string(), "Sheet1!C1:D2".to_string());

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names,
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
        };

        let rule = UnusedNamedRangesRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "PERF001");
        assert!(violations[0].message.contains("UnusedRange"));
    }
}

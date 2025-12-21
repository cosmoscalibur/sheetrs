//! PERF002: Unused sheets detection

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;
use std::collections::HashSet;

pub struct UnusedSheetsRule;

impl LinterRule for UnusedSheetsRule {
    fn id(&self) -> &str {
        "PERF002"
    }

    fn name(&self) -> &str {
        "Unused sheets"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Performance
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        // Collect all sheet names
        let all_sheets: HashSet<&str> = workbook.sheets.iter().map(|s| s.name.as_str()).collect();

        // Track which sheets are referenced
        let mut referenced_sheets = HashSet::new();

        // Check formulas for sheet references
        for sheet in &workbook.sheets {
            for cell in sheet.all_cells() {
                if let Some(formula) = cell.value.as_formula() {
                    for other_sheet in &all_sheets {
                        // Look for references like "SheetName!" in formulas
                        if formula.contains(&format!("{}!", other_sheet)) {
                            referenced_sheets.insert(*other_sheet);
                        }
                    }
                }
            }
        }

        // Check named ranges for sheet references
        for reference in workbook.defined_names.values() {
            for sheet_name in &all_sheets {
                if reference.contains(&format!("{}!", sheet_name)) {
                    referenced_sheets.insert(*sheet_name);
                }
            }
        }

        // Report sheets that are not referenced by any other sheet
        // A sheet is considered "used" if:
        // - It's the only sheet, OR
        // - It's referenced by another sheet, OR
        // - It contains formulas (it's doing work), OR
        // - We failed to parse formulas for it (safe default)
        for sheet in &workbook.sheets {
            let is_only_sheet = workbook.sheets.len() == 1;
            let is_referenced = referenced_sheets.contains(sheet.name.as_str());
            let has_formulas = sheet.cells.values().any(|c| c.value.is_formula());
            let has_content = sheet.cells.values().any(|c| !c.value.is_empty());
            let formula_error = &sheet.formula_parsing_error;

            if let Some(_err_msg) = formula_error {
                // If we failed to parse formulas, implicitly treat the sheet as "used" to avoid false positives.
                // We do NOT report this as a violation or print it, as requested by the user.
                continue;
            }

            if !is_only_sheet && !is_referenced && !has_formulas && has_content {
                violations.push(Violation::new(
                    self.id(),
                    ViolationScope::Book,
                    format!(
                        "Sheet '{}' is not referenced by any other sheet and contains no formulas",
                        sheet.name
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
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_unused_sheets() {
        let mut cells1 = HashMap::new();
        cells1.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=Sheet2!A1".to_string()),
            },
        );

        let sheet1 = Sheet {
            name: "Sheet1".to_string(),
            cells: cells1,
            used_range: Some((1, 1)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            sheet_path: None,
            formula_parsing_error: None,
        };

        let mut cells2 = HashMap::new();
        cells2.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Number(42.0),
            },
        );

        let sheet2 = Sheet {
            name: "Sheet2".to_string(),
            cells: cells2,
            used_range: Some((1, 1)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            sheet_path: None,
            formula_parsing_error: None,
        };

        let mut cells3 = HashMap::new();
        cells3.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Number(100.0),
            },
        );

        let sheet3 = Sheet {
            name: "Sheet3".to_string(),
            cells: cells3,
            used_range: Some((1, 1)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            sheet_path: None,
            formula_parsing_error: None,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet1, sheet2, sheet3],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
        };

        let rule = UnusedSheetsRule;
        let violations = rule.check(&workbook).unwrap();

        // Sheet3 should be reported as unused
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "PERF002");
        assert!(violations[0].message.contains("Sheet3"));
    }
}

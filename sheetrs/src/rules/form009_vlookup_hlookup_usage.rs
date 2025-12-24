use crate::config::LinterConfig;
use crate::reader::{CellValue, Workbook};
use crate::rules::LinterRule;
use crate::violation::{Severity, Violation, ViolationScope};

/// FORM009: VLOOKUP/HLOOKUP usage
///
/// Detects usage of VLOOKUP and HLOOKUP functions.
/// XLOOKUP or INDEX/MATCH are preferred for flexibility and performance.
pub struct VLookupHLookupUsageRule;

impl VLookupHLookupUsageRule {
    pub fn new(_config: &LinterConfig) -> Self {
        Self
    }
}

impl LinterRule for VLookupHLookupUsageRule {
    fn id(&self) -> &'static str {
        "FORM009"
    }

    fn name(&self) -> &'static str {
        "VLookupHLookupUsageRule"
    }

    fn category(&self) -> crate::rules::RuleCategory {
        crate::rules::RuleCategory::Formula
    }

    fn check(&self, workbook: &Workbook) -> anyhow::Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            for ((row, col), cell) in &sheet.cells {
                if let CellValue::Formula { formula, .. } = &cell.value {
                    let upper_formula = formula.to_uppercase();
                    if upper_formula.contains("VLOOKUP(") || upper_formula.contains("HLOOKUP(") {
                        violations.push(Violation::new(
                            self.id(),
                            ViolationScope::Cell(
                                sheet.name.clone(),
                                crate::violation::CellReference {
                                    row: *row,
                                    col: *col,
                                },
                            ),
                            "Avoid using VLOOKUP/HLOOKUP. Use XLOOKUP or INDEX/MATCH instead."
                                .to_string(),
                            Severity::Warning,
                        ));
                    }
                }
            }
        }

        Ok(violations)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::{Cell, Sheet};
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_vlookup_detection() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=VLOOKUP(A1, B:C, 2, FALSE)".to_string()),
            },
        );
        cells.insert(
            (0, 1),
            Cell {
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::formula("=HLOOKUP(A1, B:C, 2, FALSE)".to_string()),
            },
        );
        cells.insert(
            (0, 2),
            Cell {
                num_fmt: None,
                row: 0,
                col: 2,
                value: CellValue::formula("=SUM(A1:A10)".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 3)),
            hidden_columns: vec![],
            hidden_rows: vec![],
            merged_cells: vec![],
            formula_parsing_error: None,
            conditional_formatting_count: 0,
            conditional_formatting_ranges: Vec::new(),
            visible: true,
            sheet_path: None,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: vec![],
            has_macros: false,
            external_links: Vec::new(),
            external_workbooks: Vec::new(),
        };

        let config = LinterConfig::default();
        let rule = VLookupHLookupUsageRule::new(&config);

        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 2);
    }
}

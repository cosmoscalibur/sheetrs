//! ERR001: Error cell value detection

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{CellReference, Severity, Violation, ViolationScope};
use anyhow::Result;

pub struct ErrorCellsRule;

impl LinterRule for ErrorCellsRule {
    fn id(&self) -> &str {
        "ERR001"
    }

    fn name(&self) -> &str {
        "Error cell value"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::UnresolvedErrors
    }

    fn default_active(&self) -> bool {
        true
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            for cell in sheet.all_cells() {
                if cell.value.is_error() {
                    let error_value = cell.value.as_error().unwrap_or("Unknown error");
                    violations.push(Violation::new(
                        self.id(),
                        ViolationScope::Cell(
                            sheet.name.clone(),
                            CellReference::new(cell.row, cell.col),
                        ),
                        format!("Cell contains error value: {}", error_value),
                        Severity::Error,
                    ));
                }
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
    fn test_error_cells_detection() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell { num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Error("#DIV/0!".to_string()),
            },
        );
        cells.insert(
            (1, 0),
            Cell { num_fmt: None,
                row: 1,
                col: 0,
                value: CellValue::Number(42.0),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((2, 1)),
                hidden_columns: Vec::new(),
                hidden_rows: Vec::new(),
                merged_cells: Vec::new(), formula_parsing_error: None,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
        };

        let rule = ErrorCellsRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "ERR001");
        assert!(violations[0].message.contains("#DIV/0!"));
    }
}

//! PERF003: Large used range detection

use super::{LinterRule, RuleCategory};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;

pub struct LargeUsedRangeRule {
    threshold_rows: u32,
    threshold_cols: u32,
}

impl LargeUsedRangeRule {
    pub fn new(config: &LinterConfig) -> Self {
        let threshold_rows = config
            .get_rule_config("PERF003")
            .and_then(|c| c.get_int("max_extra_row"))
            .unwrap_or(2) as u32;

        let threshold_cols = config
            .get_rule_config("PERF003")
            .and_then(|c| c.get_int("max_extra_column"))
            .unwrap_or(2) as u32;

        Self {
            threshold_rows,
            threshold_cols,
        }
    }
}

impl Default for LargeUsedRangeRule {
    fn default() -> Self {
        Self {
            threshold_rows: 2,
            threshold_cols: 2,
        }
    }
}

impl LinterRule for LargeUsedRangeRule {
    fn id(&self) -> &str {
        "PERF003"
    }

    fn name(&self) -> &str {
        "Large used range detection"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Performance
    }

    fn default_active(&self) -> bool {
        true
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            if let Some((used_rows, used_cols)) = sheet.used_range {
                // Find the last cell with actual data or formula
                if let Some((last_data_row, last_data_col)) = sheet.last_data_cell() {
                    let row_diff = used_rows.saturating_sub(last_data_row + 1);
                    let col_diff = used_cols.saturating_sub(last_data_col + 1);

                    if row_diff > self.threshold_rows || col_diff > self.threshold_cols {
                        use crate::violation::CellReference;

                        let last_used_ref = CellReference::new(used_rows - 1, used_cols - 1);
                        let last_data_ref = CellReference::new(last_data_row, last_data_col);

                        violations.push(Violation::new(
                            self.id(),
                            ViolationScope::Sheet(sheet.name.clone()),
                            format!(
                                "Used range extends beyond data: last used cell {}, last data/formula cell {} (threshold: {}/{} rows/cols)",
                                last_used_ref, last_data_ref, self.threshold_rows, self.threshold_cols
                            ),
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
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_large_used_range() {
        let mut cells = HashMap::new();
        // Data only in first few cells
        cells.insert(
            (0, 0),
            Cell { num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Number(1.0),
            },
        );
        cells.insert(
            (1, 0),
            Cell { num_fmt: None,
                row: 1,
                col: 0,
                value: CellValue::Number(2.0),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            // But used range extends much further
            used_range: Some((50, 30)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(), sheet_path: None,
            formula_parsing_error: None,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
        };

        let rule = LargeUsedRangeRule::default();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "PERF003");
    }
}

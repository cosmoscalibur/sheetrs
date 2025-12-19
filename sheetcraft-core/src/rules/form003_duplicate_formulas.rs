//! PERF005: Avoid duplicate formulas

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{CellReference, Severity, Violation, ViolationScope};
use anyhow::Result;
use std::collections::{HashMap, HashSet, VecDeque};

pub struct DuplicateFormulasRule;

impl LinterRule for DuplicateFormulasRule {
    fn id(&self) -> &str {
        "FORM003"
    }

    fn name(&self) -> &str {
        "Avoid duplicate formulas"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Formula
    }

    fn default_active(&self) -> bool {
        true
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            // Group cells by formula content
            let mut formula_cells: HashMap<String, Vec<(u32, u32)>> = HashMap::new();

            for cell in sheet.all_cells() {
                if let Some(formula) = cell.value.as_formula() {
                    // Normalize formula for comparison (trim whitespace)
                    let normalized = formula.trim().to_string();
                    formula_cells
                        .entry(normalized)
                        .or_insert_with(Vec::new)
                        .push((cell.row, cell.col));
                }
            }

            // Report formulas that appear more than once
            for (formula, cells) in formula_cells {
                if cells.len() > 1 {
                    let ranges = find_contiguous_ranges(&cells);

                    // Create a single violation for this duplicated formula
                    let range_strs: Vec<String> =
                        ranges.iter().map(|r| format_single_range(r)).collect();

                    let range_list = range_strs.join(", ");

                    // Truncate formula for display if too long
                    let display_formula = if formula.chars().count() > 50 {
                        formula.chars().take(50).collect::<String>() + "..."
                    } else {
                        formula.clone()
                    };

                    violations.push(Violation::new(
                        self.id(),
                        ViolationScope::Sheet(sheet.name.clone()),
                        format!(
                            "Formula '{}' is duplicated {} times in ranges: {}. Consider using named ranges or helper cells.",
                            display_formula, cells.len(), range_list
                        ),
                        Severity::Info,
                    ));
                }
            }
        }

        Ok(violations)
    }
}

/// Format a single contiguous range
fn format_single_range(cells: &[(u32, u32)]) -> String {
    if cells.is_empty() {
        return String::new();
    }

    if cells.len() == 1 {
        return CellReference::new(cells[0].0, cells[0].1).to_string();
    }

    let min_row = cells.iter().map(|(r, _)| r).min().unwrap();
    let max_row = cells.iter().map(|(r, _)| r).max().unwrap();
    let min_col = cells.iter().map(|(_, c)| c).min().unwrap();
    let max_col = cells.iter().map(|(_, c)| c).max().unwrap();

    let start = CellReference::new(*min_row, *min_col);
    let end = CellReference::new(*max_row, *max_col);

    format!("{}:{}", start, end)
}

/// Find contiguous ranges from a list of cells
fn find_contiguous_ranges(cells: &[(u32, u32)]) -> Vec<Vec<(u32, u32)>> {
    let cell_set: HashSet<(u32, u32)> = cells.iter().copied().collect();
    let mut visited: HashSet<(u32, u32)> = HashSet::new();
    let mut ranges: Vec<Vec<(u32, u32)>> = Vec::new();

    for &cell in cells {
        if visited.contains(&cell) {
            continue;
        }

        // BFS to find all connected cells
        let mut range = Vec::new();
        let mut queue = VecDeque::new();
        queue.push_back(cell);
        visited.insert(cell);

        while let Some((row, col)) = queue.pop_front() {
            range.push((row, col));

            // Check all 4 adjacent cells (up, down, left, right)
            let neighbors = [
                (row.wrapping_sub(1), col),
                (row + 1, col),
                (row, col.wrapping_sub(1)),
                (row, col + 1),
            ];

            for neighbor in neighbors {
                if cell_set.contains(&neighbor) && !visited.contains(&neighbor) {
                    visited.insert(neighbor);
                    queue.push_back(neighbor);
                }
            }
        }

        ranges.push(range);
    }

    ranges
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_duplicate_formulas() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Formula("=A1+B1".to_string()),
            },
        );
        cells.insert(
            (1, 0),
            Cell {
                num_fmt: None,
                row: 1,
                col: 0,
                value: CellValue::Formula("=A1+B1".to_string()),
            },
        );
        cells.insert(
            (2, 0),
            Cell {
                num_fmt: None,
                row: 2,
                col: 0,
                value: CellValue::Formula("=A1+B1".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((3, 1)),
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
        };

        let rule = DuplicateFormulasRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "FORM003");
        assert!(violations[0].message.contains("duplicated 3 times"));
    }

    #[test]
    fn test_unique_formulas() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Formula("=A1+B1".to_string()),
            },
        );
        cells.insert(
            (1, 0),
            Cell {
                num_fmt: None,
                row: 1,
                col: 0,
                value: CellValue::Formula("=A2+B2".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((2, 1)),
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
        };

        let rule = DuplicateFormulasRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 0);
    }
}

//! FORM003: Deep formula nesting detection

use super::{LinterRule, RuleCategory};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::violation::{CellReference, Severity, Violation, ViolationScope};
use anyhow::Result;
use std::collections::{HashSet, VecDeque};

pub struct DeepFormulaNestingRule {
    max_nesting: usize,
}

impl DeepFormulaNestingRule {
    pub fn new(config: &LinterConfig) -> Self {
        let max_nesting = config
            .get_param_int("max_formula_nesting", None)
            .unwrap_or(5) as usize;

        Self { max_nesting }
    }
}

impl Default for DeepFormulaNestingRule {
    fn default() -> Self {
        Self { max_nesting: 5 }
    }
}

impl LinterRule for DeepFormulaNestingRule {
    fn id(&self) -> &str {
        "FORM006"
    }

    fn name(&self) -> &str {
        "Deep formula nesting"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Formula
    }

    fn default_active(&self) -> bool {
        false
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            let mut deep_nesting_cells: Vec<(u32, u32)> = Vec::new();

            for cell in sheet.all_cells() {
                if let Some(formula) = cell.value.as_formula() {
                    let nesting_depth = calculate_nesting_depth(formula);

                    if nesting_depth > self.max_nesting {
                        deep_nesting_cells.push((cell.row, cell.col));
                    }
                }
            }

            // Group into contiguous ranges and create violations
            if !deep_nesting_cells.is_empty() {
                let ranges = find_contiguous_ranges(&deep_nesting_cells);

                for range in ranges {
                    let range_str = format_single_range(&range);
                    violations.push(Violation::new(
                        self.id(),
                        ViolationScope::Sheet(sheet.name.clone()),
                        format!(
                            "Formula with deep nesting (>{} levels) in range: {}. Consider simplifying.",
                            self.max_nesting, range_str
                        ),
                        Severity::Warning,
                    ));
                }
            }
        }

        Ok(violations)
    }
}

/// Calculate the maximum nesting depth of a formula
/// Uses parenthesis counting to determine depth
fn calculate_nesting_depth(formula: &str) -> usize {
    let mut max_depth: usize = 0;
    let mut current_depth: usize = 0;

    for ch in formula.chars() {
        match ch {
            '(' => {
                current_depth += 1;
                max_depth = max_depth.max(current_depth);
            }
            ')' => {
                current_depth = current_depth.saturating_sub(1);
            }
            _ => {}
        }
    }

    max_depth
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
    fn test_deep_nesting() {
        let mut cells = HashMap::new();
        // Formula with 6 levels of nesting
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Formula(
                    "=IF(A1,IF(B1,IF(C1,IF(D1,IF(E1,IF(F1,1,0),0),0),0),0),0)".to_string(),
                ),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 1)),
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

        let rule = DeepFormulaNestingRule::default();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "FORM006");
        assert!(violations[0].message.contains(">5 levels"));
    }

    #[test]
    fn test_shallow_nesting() {
        let mut cells = HashMap::new();
        // Formula with 3 levels of nesting
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Formula("=IF(A1,IF(B1,IF(C1,1,0),0),0)".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 1)),
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

        let rule = DeepFormulaNestingRule::default();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_nesting_depth_calculation() {
        assert_eq!(calculate_nesting_depth("=A1+B1"), 0);
        assert_eq!(calculate_nesting_depth("=SUM(A1:A10)"), 1);
        assert_eq!(calculate_nesting_depth("=IF(A1,SUM(B1:B10),0)"), 2);
        assert_eq!(calculate_nesting_depth("=IF(A1,IF(B1,SUM(C1:C10),0),0)"), 3);
    }
}

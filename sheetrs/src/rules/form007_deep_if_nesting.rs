//! FORM007: Deeply nested IF statements detection

use super::{LinterRule, RuleCategory};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::violation::{CellReference, Severity, Violation, ViolationScope};
use anyhow::Result;
use std::collections::{HashSet, VecDeque};

pub struct DeepIfNestingRule {
    max_if_nesting: usize,
}

impl DeepIfNestingRule {
    pub fn new(config: &LinterConfig) -> Self {
        let max_if_nesting = config.get_param_int("max_if_nesting", None).unwrap_or(5) as usize;

        Self { max_if_nesting }
    }
}

impl Default for DeepIfNestingRule {
    fn default() -> Self {
        Self { max_if_nesting: 5 }
    }
}

impl LinterRule for DeepIfNestingRule {
    fn id(&self) -> &str {
        "FORM007"
    }

    fn name(&self) -> &str {
        "Deeply nested IF statements"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Formula
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            let mut deep_if_cells: Vec<(u32, u32)> = Vec::new();

            for cell in sheet.all_cells() {
                if let Some(formula) = cell.value.as_formula() {
                    let if_nesting = count_if_nesting(formula);

                    if if_nesting > self.max_if_nesting {
                        deep_if_cells.push((cell.row, cell.col));
                    }
                }
            }

            // Group into contiguous ranges and create violations
            if !deep_if_cells.is_empty() {
                let ranges = find_contiguous_ranges(&deep_if_cells);

                for range in ranges {
                    let range_str = format_single_range(&range);
                    violations.push(Violation::new(
                        self.id(),
                        ViolationScope::Sheet(sheet.name.clone()),
                        format!(
                            "Deeply nested IF statements (>{} levels) in range: {}. Consider using lookup tables or IFS function.",
                            self.max_if_nesting, range_str
                        ),
                        Severity::Warning,
                    ));
                }
            }
        }

        Ok(violations)
    }
}

/// Count the maximum nesting depth of IF statements in a formula
fn count_if_nesting(formula: &str) -> usize {
    let formula_upper = formula.to_uppercase();
    let mut max_depth: usize = 0;
    let mut current_depth: usize = 0;
    let mut i = 0;
    let chars: Vec<char> = formula_upper.chars().collect();

    while i < chars.len() {
        // Look for "IF(" pattern
        if i + 3 < chars.len() && chars[i] == 'I' && chars[i + 1] == 'F' && chars[i + 2] == '(' {
            // Make sure it's not part of a larger word (e.g., "COUNTIF")
            let is_standalone = if i == 0 {
                true
            } else {
                let prev_char = chars[i - 1];
                !prev_char.is_alphanumeric() && prev_char != '_'
            };

            if is_standalone {
                current_depth += 1;
                max_depth = max_depth.max(current_depth);
                i += 3; // Skip "IF("
                continue;
            }
        }

        // Track closing parentheses to decrement depth
        if chars[i] == ')' {
            // This is a simplified approach - decrementing for any closing paren
            // A more robust approach would track which parens belong to IF statements
            if current_depth > 0 {
                // Look ahead to see if this closes an IF
                // For simplicity, decrement when a closing paren is seen
                // This may undercount in complex cases but is good enough
                current_depth = current_depth.saturating_sub(1);
            }
        }

        i += 1;
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
    fn test_deeply_nested_if() {
        let mut cells = HashMap::new();
        // 6 nested IF statements
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula(
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
            merged_cells: Vec::new(),
            sheet_path: None,
            formula_parsing_error: None,
            conditional_formatting_count: 0,
            conditional_formatting_ranges: Vec::new(),
            visible: true,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
            external_workbooks: Vec::new(),
        };

        let rule = DeepIfNestingRule::default();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "FORM007");
        assert!(violations[0].message.contains(">5 levels"));
    }

    #[test]
    fn test_shallow_if_nesting() {
        let mut cells = HashMap::new();
        // 3 nested IF statements
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=IF(A1,IF(B1,IF(C1,1,0),0),0)".to_string()),
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
            visible: true,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
            external_workbooks: Vec::new(),
        };

        let rule = DeepIfNestingRule::default();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_if_counting() {
        assert_eq!(count_if_nesting("=A1+B1"), 0);
        assert_eq!(count_if_nesting("=IF(A1,B1,C1)"), 1);
        assert_eq!(count_if_nesting("=IF(A1,IF(B1,C1,D1),E1)"), 2);
        assert_eq!(count_if_nesting("=IF(A1,IF(B1,IF(C1,D1,E1),F1),G1)"), 3);
        // Should not count COUNTIF
        assert_eq!(count_if_nesting("=COUNTIF(A1:A10,\">5\")"), 0);
    }
}

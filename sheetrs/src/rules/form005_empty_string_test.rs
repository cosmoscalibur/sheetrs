//! FORM002: Empty string test → ISBLANK

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{CellReference, Severity, Violation, ViolationScope};
use anyhow::Result;
use regex::Regex;
use std::collections::{HashSet, VecDeque};

pub struct EmptyStringTestRule {
    // Regex patterns for detecting empty string tests
    patterns: Vec<(Regex, &'static str)>,
}

impl EmptyStringTestRule {
    pub fn new() -> Self {
        let patterns = vec![
            // Pattern: ="" or <>""
            (Regex::new(r#"(=|<>)\s*"""#).unwrap(), r#"=""#),
            // Pattern: LEN(...)=0 or LEN(...)>0
            (
                Regex::new(r"LEN\s*\([^)]+\)\s*(=|<>|>|<)\s*0").unwrap(),
                "LEN(...)=0",
            ),
        ];

        Self { patterns }
    }
}

impl Default for EmptyStringTestRule {
    fn default() -> Self {
        Self::new()
    }
}

impl LinterRule for EmptyStringTestRule {
    fn id(&self) -> &str {
        "FORM005"
    }

    fn name(&self) -> &str {
        "Empty string test → ISBLANK"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Formula
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            let mut empty_test_cells: Vec<(u32, u32)> = Vec::new();

            for cell in sheet.all_cells() {
                if let Some(formula) = cell.value.as_formula() {
                    let formula_upper = formula.to_uppercase();

                    // Check if any pattern matches
                    let has_empty_test = self
                        .patterns
                        .iter()
                        .any(|(pattern, _)| pattern.is_match(&formula_upper));

                    if has_empty_test {
                        empty_test_cells.push((cell.row, cell.col));
                    }
                }
            }

            // Report empty string tests
            if !empty_test_cells.is_empty() {
                let ranges = find_contiguous_ranges(&empty_test_cells);

                for range in ranges {
                    let range_str = format_single_range(&range);
                    violations.push(Violation::new(
                        self.id(),
                        ViolationScope::Sheet(sheet.name.clone()),
                        format!(
                            "Empty string test (=\"\" or LEN()=0) found in range: {}. Consider using ISBLANK() for better readability.",
                            range_str
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
    fn test_empty_string_equals() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula(r#"=IF(A1="","Empty","Not Empty")"#.to_string()),
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

        let rule = EmptyStringTestRule::new();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "FORM005");
        assert!(violations[0].message.contains("ISBLANK"));
    }

    #[test]
    fn test_empty_string_not_equals() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula(r#"=IF(A1<>"","Not Empty","Empty")"#.to_string()),
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

        let rule = EmptyStringTestRule::new();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "FORM005");
    }

    #[test]
    fn test_len_equals_zero() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=IF(LEN(A1)=0,\"Empty\",\"Not Empty\")".to_string()),
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

        let rule = EmptyStringTestRule::new();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "FORM005");
    }

    #[test]
    fn test_isblank_usage() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=IF(ISBLANK(A1),\"Empty\",\"Not Empty\")".to_string()),
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

        let rule = EmptyStringTestRule::new();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 0);
    }
}

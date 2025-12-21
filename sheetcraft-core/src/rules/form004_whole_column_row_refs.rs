//! FORM001: Avoid whole-column or whole-row references

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{CellReference, Severity, Violation, ViolationScope};
use anyhow::Result;
use regex::Regex;
use std::collections::{HashSet, VecDeque};

pub struct WholeColumnRowRefsRule {
    // Regex patterns for detecting whole column/row references
    column_pattern: Regex,
    row_pattern: Regex,
}

impl WholeColumnRowRefsRule {
    pub fn new() -> Self {
        // Pattern for whole column references: A:A, A:Z, etc.
        // Matches one or more letters, colon, one or more letters
        let column_pattern = Regex::new(r"\b[A-Z]+:[A-Z]+\b").unwrap();

        // Pattern for whole row references: 1:1, 1:100, etc.
        // Matches one or more digits, colon, one or more digits
        let row_pattern = Regex::new(r"\b\d+:\d+\b").unwrap();

        Self {
            column_pattern,
            row_pattern,
        }
    }
}

impl Default for WholeColumnRowRefsRule {
    fn default() -> Self {
        Self::new()
    }
}

impl LinterRule for WholeColumnRowRefsRule {
    fn id(&self) -> &str {
        "FORM004"
    }

    fn name(&self) -> &str {
        "Avoid whole-column or whole-row references"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Formula
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            let mut column_ref_cells: Vec<(u32, u32)> = Vec::new();
            let mut row_ref_cells: Vec<(u32, u32)> = Vec::new();

            for cell in sheet.all_cells() {
                if let Some(formula) = cell.value.as_formula() {
                    let formula_upper = formula.to_uppercase();

                    let has_column_ref = self.column_pattern.is_match(&formula_upper);
                    let has_row_ref = self.row_pattern.is_match(&formula_upper);

                    if has_column_ref {
                        column_ref_cells.push((cell.row, cell.col));
                    } else if has_row_ref {
                        row_ref_cells.push((cell.row, cell.col));
                    }
                }
            }

            // Report whole column references
            if !column_ref_cells.is_empty() {
                let ranges = find_contiguous_ranges(&column_ref_cells);

                for range in ranges {
                    let range_str = format_single_range(&range);
                    violations.push(Violation::new(
                        self.id(),
                        ViolationScope::Sheet(sheet.name.clone()),
                        format!(
                            "Whole-column reference (e.g., A:A) found in range: {}. Use bounded ranges for better performance.",
                            range_str
                        ),
                        Severity::Warning,
                    ));
                }
            }

            // Report whole row references
            if !row_ref_cells.is_empty() {
                let ranges = find_contiguous_ranges(&row_ref_cells);

                for range in ranges {
                    let range_str = format_single_range(&range);
                    violations.push(Violation::new(
                        self.id(),
                        ViolationScope::Sheet(sheet.name.clone()),
                        format!(
                            "Whole-row reference (e.g., 1:1) found in range: {}. Use bounded ranges for better performance.",
                            range_str
                        ),
                        Severity::Warning,
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
    fn test_whole_column_reference() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=SUM(A:A)".to_string()),
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

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
        };

        let rule = WholeColumnRowRefsRule::new();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "FORM004");
        assert!(violations[0].message.contains("Whole-column"));
    }

    #[test]
    fn test_whole_row_reference() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=SUM(1:1)".to_string()),
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

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
        };

        let rule = WholeColumnRowRefsRule::new();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "FORM004");
        assert!(violations[0].message.contains("Whole-row"));
    }

    #[test]
    fn test_bounded_reference() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=SUM(A1:A10)".to_string()),
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

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
        };

        let rule = WholeColumnRowRefsRule::new();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 0);
    }
}

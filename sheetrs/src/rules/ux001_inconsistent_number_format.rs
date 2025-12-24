//! UX001: Inconsistent number formatting detection
//! Detects numeric data stored as text instead of as number type

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;

pub struct NumericTextRule;

impl LinterRule for NumericTextRule {
    fn id(&self) -> &str {
        "UX001"
    }

    fn name(&self) -> &str {
        "Numeric data stored as text"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::FormattingAndUsability
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            // Collect all cells with numeric text
            let mut numeric_text_cells: Vec<(u32, u32)> = Vec::new();

            for cell in sheet.all_cells() {
                // Check if cell contains text that looks like a number
                if let crate::reader::workbook::CellValue::Text(text) = &cell.value {
                    if is_numeric_text(text) {
                        numeric_text_cells.push((cell.row, cell.col));
                    }
                }
            }

            // Group cells into ranges and create violations
            if !numeric_text_cells.is_empty() {
                let ranges = find_contiguous_ranges(&numeric_text_cells);

                // Create a separate violation for each contiguous range
                for range in ranges {
                    let range_str = format_single_range(&range);
                    violations.push(Violation::new(
                        self.id(),
                        ViolationScope::Sheet(sheet.name.clone()),
                        format!("Numeric data stored as text in range: {}", range_str),
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
    use crate::violation::CellReference;

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
    use std::collections::{HashSet, VecDeque};

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

/// Check if a text string represents a numeric value
fn is_numeric_text(text: &str) -> bool {
    let trimmed = text.trim();

    // Empty strings are not numeric
    if trimmed.is_empty() {
        return false;
    }

    // Try to parse as a number
    // This handles integers, floats, scientific notation, etc.
    trimmed.parse::<f64>().is_ok()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_numeric_text_detection() {
        let mut cells = HashMap::new();

        // Numeric text - should be flagged
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text("42".to_string()),
            },
        );

        cells.insert(
            (1, 0),
            Cell {
                num_fmt: None,
                row: 1,
                col: 0,
                value: CellValue::Text("3.14".to_string()),
            },
        );

        // Actual number - should NOT be flagged
        cells.insert(
            (2, 0),
            Cell {
                num_fmt: None,
                row: 2,
                col: 0,
                value: CellValue::Number(100.0),
            },
        );

        // Non-numeric text - should NOT be flagged
        cells.insert(
            (3, 0),
            Cell {
                num_fmt: None,
                row: 3,
                col: 0,
                value: CellValue::Text("Hello".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((4, 1)),
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

        let rule = NumericTextRule;
        let violations = rule.check(&workbook).unwrap();

        // Should detect numeric text as a range
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "UX001");
        assert!(violations[0].message.contains("range"));
        // The range should be "2 cells in A1:B2" since cells are at (0,0) and (1,0)
    }
}

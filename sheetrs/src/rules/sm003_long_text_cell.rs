//! SM003: Long text cell

use super::{LinterRule, RuleCategory};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;
use std::collections::{HashSet, VecDeque};

pub struct LongTextCellRule {
    config: LinterConfig,
}

impl LongTextCellRule {
    pub fn new(config: &LinterConfig) -> Self {
        Self {
            config: config.clone(),
        }
    }
}

impl Default for LongTextCellRule {
    fn default() -> Self {
        Self {
            config: LinterConfig::default(),
        }
    }
}

impl LinterRule for LongTextCellRule {
    fn id(&self) -> &str {
        "SM003"
    }

    fn name(&self) -> &str {
        "Long text cell"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::StructuralAndMaintainability
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            let threshold = self
                .config
                .get_param_int("max_text_length", Some(&sheet.name))
                .unwrap_or(255) as usize;

            let mut long_text_cells: Vec<(u32, u32)> = Vec::new();

            for cell in sheet.all_cells() {
                if let crate::reader::workbook::CellValue::Text(text) = &cell.value {
                    if text.len() > threshold {
                        long_text_cells.push((cell.row, cell.col));
                    }
                }
            }

            // Group into contiguous ranges and create violations
            if !long_text_cells.is_empty() {
                let ranges = find_contiguous_ranges(&long_text_cells);

                for range in ranges {
                    let range_str = format_single_range(&range);
                    violations.push(Violation::new(
                        self.id(),
                        ViolationScope::Sheet(sheet.name.clone()),
                        format!(
                            "Long text cells (>{} characters) in range: {}",
                            threshold, range_str
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
    fn test_long_text_cell() {
        let mut cells = HashMap::new();
        let long_text = "A".repeat(300); // >255 chars

        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text(long_text),
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

        let rule = LongTextCellRule::default();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "SM003");
        assert!(violations[0].message.contains("range"));
    }
}

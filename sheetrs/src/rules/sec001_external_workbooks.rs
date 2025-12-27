//! SEC001: External workbook references

use super::{LinterRule, RuleCategory};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkScope {
    Book,  // Only report book-level violations
    Sheet, // Only report sheet-level violations
}

impl LinkScope {
    fn from_str(s: &str) -> Self {
        match s.to_uppercase().as_str() {
            "BOOK" => LinkScope::Book,
            "SHEET" => LinkScope::Sheet,
            _ => LinkScope::Book, // Default: BOOK
        }
    }
}

pub struct ExternalWorkbooksRule {
    scope: LinkScope,
}

impl ExternalWorkbooksRule {
    pub fn new(config: &LinterConfig) -> Self {
        let scope = config
            .get_param_str("external_workbook_scope", None)
            .map(LinkScope::from_str)
            .unwrap_or(LinkScope::Book);
        Self { scope }
    }
}

impl Default for ExternalWorkbooksRule {
    fn default() -> Self {
        Self {
            scope: LinkScope::Book,
        }
    }
}

impl LinterRule for ExternalWorkbooksRule {
    fn id(&self) -> &str {
        "SEC001"
    }

    fn name(&self) -> &str {
        "External workbook references"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::SecurityAndPrivacy
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        // BOOK scope: from external_workbooks field
        if matches!(self.scope, LinkScope::Book) {
            for wb in &workbook.external_workbooks {
                violations.push(Violation::new(
                    self.id(),
                    ViolationScope::Book,
                    format!("External workbook '{}' found in metadata.", wb.path),
                    Severity::Warning,
                ));
            }
        }

        // SHEET scope: from formulas
        if matches!(self.scope, LinkScope::Sheet) {
            for sheet in &workbook.sheets {
                let mut workbook_cells: Vec<(u32, u32, usize)> = Vec::new();

                for cell in sheet.all_cells() {
                    if let Some(formula) = cell.value.as_formula() {
                        let indices = extract_external_workbook_indices(formula);
                        for idx in indices {
                            workbook_cells.push((cell.row, cell.col, idx));
                        }
                    }
                }

                let grouped = group_cells_by_index(workbook_cells);
                for (idx, cells) in grouped {
                    let wb_name = workbook
                        .external_workbooks
                        .get(idx)
                        .map(|wb| wb.path.as_str())
                        .unwrap_or("unknown");

                    let ranges = find_contiguous_ranges(&cells);
                    for range in ranges {
                        violations.push(Violation::new(
                            self.id(),
                            ViolationScope::Sheet(sheet.name.clone()),
                            format!(
                                "External workbook reference {} found in range: {}",
                                wb_name,
                                format_single_range(&range)
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

/// Extract [N] indices from formula
fn extract_external_workbook_indices(formula: &str) -> Vec<usize> {
    use regex::Regex;
    use std::sync::OnceLock;

    static INDEX_PATTERN: OnceLock<Regex> = OnceLock::new();
    let re = INDEX_PATTERN.get_or_init(|| Regex::new(r"\[(\d+)\]").unwrap());

    let mut indices = Vec::new();
    for cap in re.captures_iter(formula) {
        if let Some(num_str) = cap.get(1) {
            if let Ok(num) = num_str.as_str().parse::<usize>() {
                if num > 0 {
                    indices.push(num - 1); // Convert 1-based to 0-based
                }
            }
        }
    }
    indices
}

/// Group cells by workbook index
///
/// Groups cells that reference the same external workbook to create consolidated violations.
///
/// Example:
///   A1: =[1]Sheet1!B2  (workbook index 1 = test.xlsx)
///   A2: =[1]Sheet1!C3  (workbook index 1 = test.xlsx)
///   A3: =[2]Sheet1!D4  (workbook index 2 = other.xlsx)
///
/// Without grouping: 3 violations (one per cell)
/// With grouping: 2 violations (one per workbook)
///   - "test.xlsx found in range: A1:A2"
///   - "other.xlsx found in range: A3"
///
/// This makes output cleaner and shows which workbooks are used where.
fn group_cells_by_index(cells: Vec<(u32, u32, usize)>) -> Vec<(usize, Vec<(u32, u32)>)> {
    use std::collections::HashMap;

    let mut grouped: HashMap<usize, Vec<(u32, u32)>> = HashMap::new();
    for (row, col, idx) in cells {
        grouped.entry(idx).or_default().push((row, col));
    }
    grouped.into_iter().collect()
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_external_workbook_in_formula() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=[1]Sheet1!A1".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_workbooks: vec![crate::reader::ExternalWorkbook {
                index: 0,
                path: "Book1.xlsx".to_string(),
            }],
        };

        // Test SHEET scope
        let rule = ExternalWorkbooksRule {
            scope: LinkScope::Sheet,
        };
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "SEC001");
        assert!(violations[0].message.contains("Book1.xlsx"));
        assert!(violations[0].message.contains("range"));
    }

    #[test]
    fn test_external_workbook_in_metadata() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells: HashMap::new(),
            used_range: Some((0, 0)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_workbooks: vec![crate::reader::ExternalWorkbook {
                index: 0,
                path: "external_workbook.xlsx".to_string(),
            }],
        };

        // Test BOOK scope (default)
        let rule = ExternalWorkbooksRule::default();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].scope, ViolationScope::Book);
        assert!(violations[0].message.contains("external_workbook.xlsx"));
        assert!(violations[0].message.contains("metadata"));
    }
}

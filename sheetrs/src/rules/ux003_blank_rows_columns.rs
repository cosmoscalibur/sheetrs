//! UX004: Blank rows/columns in used ranges

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;

pub struct BlankRowsColumnsRule {
    max_blank_row: u32,
    max_blank_column: u32,
}

impl BlankRowsColumnsRule {
    pub fn new(config: &crate::config::LinterConfig) -> Self {
        let max_blank_row = config
            .get_param_int("max_blank_row", Some("UX004"))
            .unwrap_or(2) as u32;
        let max_blank_column = config
            .get_param_int("max_blank_column", Some("UX004"))
            .unwrap_or(2) as u32;

        Self {
            max_blank_row,
            max_blank_column,
        }
    }
}

impl LinterRule for BlankRowsColumnsRule {
    fn id(&self) -> &str {
        "UX003"
    }

    fn name(&self) -> &str {
        "Blank rows/columns in used ranges"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::FormattingAndUsability
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            // Skip sheets with no data
            if sheet.cells.is_empty() {
                continue;
            }

            // Find the used range
            let (min_row, max_row, min_col, max_col) = find_used_range(sheet);

            // Check for blank rows within used range
            let blank_rows = find_blank_rows(sheet, min_row, max_row, min_col, max_col);
            if !blank_rows.is_empty() {
                // Group contiguous rows and filter by max_blank_row
                let contiguous_groups = group_contiguous_indices(&blank_rows);
                let filtered_groups: Vec<Vec<u32>> = contiguous_groups
                    .into_iter()
                    .filter(|group| group.len() as u32 > self.max_blank_row)
                    .collect();

                if !filtered_groups.is_empty() {
                    // Flatten groups to format ranges
                    let mut all_filtered_rows = Vec::new();
                    for group in filtered_groups {
                        all_filtered_rows.extend(group);
                    }

                    let ranges = format_row_ranges(&all_filtered_rows);
                    violations.push(Violation::new(
                        self.id(),
                        ViolationScope::Sheet(sheet.name.clone()),
                        format!(
                            "Blank rows within used range: {}. Consider removing or filling these rows.",
                            ranges
                        ),
                        Severity::Info,
                    ));
                }
            }

            // Check for blank columns within used range
            let blank_cols = find_blank_columns(sheet, min_row, max_row, min_col, max_col);
            if !blank_cols.is_empty() {
                // Group contiguous columns and filter by max_blank_column
                let contiguous_groups = group_contiguous_indices(&blank_cols);
                let filtered_groups: Vec<Vec<u32>> = contiguous_groups
                    .into_iter()
                    .filter(|group| group.len() as u32 > self.max_blank_column)
                    .collect();

                if !filtered_groups.is_empty() {
                    let mut all_filtered_cols = Vec::new();
                    for group in filtered_groups {
                        all_filtered_cols.extend(group);
                    }

                    let ranges = format_column_ranges(&all_filtered_cols);
                    violations.push(Violation::new(
                        self.id(),
                        ViolationScope::Sheet(sheet.name.clone()),
                        format!(
                            "Blank columns within used range: {}. Consider removing or filling these columns.",
                            ranges
                        ),
                        Severity::Info,
                    ));
                }
            }
        }

        Ok(violations)
    }
}

/// Group contiguous indices into ranges (Helper from SEC003 logic)
fn group_contiguous_indices(indices: &[u32]) -> Vec<Vec<u32>> {
    if indices.is_empty() {
        return Vec::new();
    }

    let mut sorted = indices.to_vec();
    sorted.sort_unstable();
    sorted.dedup();

    let mut ranges = Vec::new();
    let mut current_range = vec![sorted[0]];

    for &idx in &sorted[1..] {
        if idx == *current_range.last().unwrap() + 1 {
            current_range.push(idx);
        } else {
            ranges.push(current_range.clone());
            current_range = vec![idx];
        }
    }
    ranges.push(current_range);

    ranges
}

/// Find the used range (min/max row and column with data)
fn find_used_range(sheet: &crate::reader::workbook::Sheet) -> (u32, u32, u32, u32) {
    let mut min_row = u32::MAX;
    let mut max_row = 0u32;
    let mut min_col = u32::MAX;
    let mut max_col = 0u32;

    for cell in sheet.all_cells() {
        min_row = min_row.min(cell.row);
        max_row = max_row.max(cell.row);
        min_col = min_col.min(cell.col);
        max_col = max_col.max(cell.col);
    }

    (min_row, max_row, min_col, max_col)
}

/// Find blank rows within the used range
fn find_blank_rows(
    sheet: &crate::reader::workbook::Sheet,
    min_row: u32,
    max_row: u32,
    min_col: u32,
    max_col: u32,
) -> Vec<u32> {
    let mut blank_rows = Vec::new();

    for row in min_row..=max_row {
        let has_data = (min_col..=max_col).any(|col| sheet.cells.contains_key(&(row, col)));

        if !has_data {
            blank_rows.push(row);
        }
    }

    blank_rows
}

/// Find blank columns within the used range
fn find_blank_columns(
    sheet: &crate::reader::workbook::Sheet,
    min_row: u32,
    max_row: u32,
    min_col: u32,
    max_col: u32,
) -> Vec<u32> {
    let mut blank_cols = Vec::new();

    for col in min_col..=max_col {
        let has_data = (min_row..=max_row).any(|row| sheet.cells.contains_key(&(row, col)));

        if !has_data {
            blank_cols.push(col);
        }
    }

    blank_cols
}

/// Format row ranges (e.g., "3, 5-7, 10")
fn format_row_ranges(rows: &[u32]) -> String {
    format_ranges(rows, |r| (r + 1).to_string()) // Convert to 1-based
}

/// Format column ranges (e.g., "C, E-G, J")
fn format_column_ranges(cols: &[u32]) -> String {
    format_ranges(cols, column_index_to_letter)
}

/// Generic range formatter
fn format_ranges<F>(indices: &[u32], formatter: F) -> String
where
    F: Fn(u32) -> String,
{
    if indices.is_empty() {
        return String::new();
    }

    let mut ranges = Vec::new();
    let mut start = indices[0];
    let mut end = indices[0];

    for &idx in &indices[1..] {
        if idx == end + 1 {
            end = idx;
        } else {
            if start == end {
                ranges.push(formatter(start));
            } else {
                ranges.push(format!("{}-{}", formatter(start), formatter(end)));
            }
            start = idx;
            end = idx;
        }
    }

    // Add the last range
    if start == end {
        ranges.push(formatter(start));
    } else {
        ranges.push(format!("{}-{}", formatter(start), formatter(end)));
    }

    ranges.join(", ")
}

/// Convert column index (0-based) to Excel-style letter (A, B, ..., Z, AA, AB, ...)
fn column_index_to_letter(col: u32) -> String {
    let mut result = String::new();
    let mut col = col + 1; // Convert to 1-based

    while col > 0 {
        col -= 1;
        result.insert(0, (b'A' + (col % 26) as u8) as char);
        col /= 26;
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_blank_rows() {
        let mut cells = HashMap::new();
        // Row 0: A1, B1
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text("A1".to_string()),
            },
        );
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text("A1".to_string()),
            },
        );
        cells.insert(
            (0, 1),
            Cell {
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Text("B1".to_string()),
            },
        );
        // Row 1: blank -> 1 contiguous
        // Row 2: A3, B3
        cells.insert(
            (2, 0),
            Cell {
                num_fmt: None,
                row: 2,
                col: 0,
                value: CellValue::Text("A3".to_string()),
            },
        );
        cells.insert(
            (2, 1),
            Cell {
                num_fmt: None,
                row: 2,
                col: 1,
                value: CellValue::Text("B3".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((3, 2)),
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

        // Limit 0: should catch 1 blank row
        let rule = BlankRowsColumnsRule {
            max_blank_row: 0,
            max_blank_column: 0,
        };
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "UX003");
        assert!(violations[0].message.contains("Blank rows"));
        assert!(violations[0].message.contains("2")); // Row 2 (1-based)

        // Limit 1: should NOT catch 1 blank row
        let rule_relaxed = BlankRowsColumnsRule {
            max_blank_row: 1,
            max_blank_column: 1,
        };
        let violations_relaxed = rule_relaxed.check(&workbook).unwrap();
        assert_eq!(violations_relaxed.len(), 0);
    }

    #[test]
    fn test_blank_columns() {
        let mut cells = HashMap::new();
        // Column A: A1, A2
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text("A1".to_string()),
            },
        );
        cells.insert(
            (1, 0),
            Cell {
                num_fmt: None,
                row: 1,
                col: 0,
                value: CellValue::Text("A2".to_string()),
            },
        );
        // Column B: blank -> 1 contiguous
        // Column C: C1, C2
        cells.insert(
            (0, 2),
            Cell {
                num_fmt: None,
                row: 0,
                col: 2,
                value: CellValue::Text("C1".to_string()),
            },
        );
        cells.insert(
            (1, 2),
            Cell {
                num_fmt: None,
                row: 1,
                col: 2,
                value: CellValue::Text("C2".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((2, 3)),
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

        // Limit 0: should catch 1 blank column
        let rule = BlankRowsColumnsRule {
            max_blank_row: 0,
            max_blank_column: 0,
        };
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "UX003");
        assert!(violations[0].message.contains("Blank columns"));
        assert!(violations[0].message.contains("B")); // Column B

        // Limit 1: should NOT catch 1 blank column
        let rule_relaxed = BlankRowsColumnsRule {
            max_blank_row: 1,
            max_blank_column: 1,
        };
        let violations_relaxed = rule_relaxed.check(&workbook).unwrap();
        assert_eq!(violations_relaxed.len(), 0);
    }

    #[test]
    fn test_no_blanks() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text("A1".to_string()),
            },
        );
        cells.insert(
            (0, 1),
            Cell {
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Text("B1".to_string()),
            },
        );
        cells.insert(
            (1, 0),
            Cell {
                num_fmt: None,
                row: 1,
                col: 0,
                value: CellValue::Text("A2".to_string()),
            },
        );
        cells.insert(
            (1, 1),
            Cell {
                num_fmt: None,
                row: 1,
                col: 1,
                value: CellValue::Text("B2".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((2, 2)),
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

        let rule = BlankRowsColumnsRule {
            max_blank_row: 2,
            max_blank_column: 2,
        };
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_column_index_to_letter() {
        assert_eq!(column_index_to_letter(0), "A");
        assert_eq!(column_index_to_letter(25), "Z");
        assert_eq!(column_index_to_letter(26), "AA");
        assert_eq!(column_index_to_letter(27), "AB");
    }
}

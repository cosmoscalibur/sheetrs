//! UX003: Blank rows/columns in used ranges

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
        let max_blank_row = config.get_param_int("max_blank_row", None).unwrap_or(2) as u32;
        let max_blank_column = config.get_param_int("max_blank_column", None).unwrap_or(2) as u32;

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
            // Skip sheets with no cells to avoid noise and potential overflows
            if sheet.cells.is_empty() {
                continue;
            }

            // Use sheet.used_range metadata instead of recalculating
            // This ensures we include styled cells in the range
            let (min_row, max_row, min_col, max_col) =
                if let Some((used_rows, used_cols)) = sheet.used_range {
                    // used_range is in count format (1-indexed max + 1), convert to 0-indexed positions
                    // For PERF003 display, we subtract 1. Here we need actual positions.
                    // Actually, used_range stores (max_row+1, max_col+1) so we need to subtract 1
                    let max_row = used_rows.saturating_sub(1);
                    let max_col = used_cols.saturating_sub(1);

                    // Find min from actual cells
                    let (cell_min_row, cell_min_col) = if sheet.cells.is_empty() {
                        (0, 0)
                    } else {
                        sheet
                            .cells
                            .keys()
                            .fold((u32::MAX, u32::MAX), |(min_r, min_c), (r, c)| {
                                (min_r.min(*r), min_c.min(*c))
                            })
                    };

                    (cell_min_row, max_row, cell_min_col, max_col)
                } else {
                    // Fallback to calculating from cells if no used_range metadata
                    find_used_range(sheet)
                };

            // Check for blank rows/columns BEFORE used range (from row/col 0)
            if min_row > 0 {
                let blank_rows_before: Vec<u32> = (0..min_row).collect();
                if !blank_rows_before.is_empty()
                    && blank_rows_before.len() as u32 > self.max_blank_row
                {
                    let ranges = format_row_ranges(&blank_rows_before);
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

            if min_col > 0 {
                let blank_cols_before: Vec<u32> = (0..min_col).collect();
                if !blank_cols_before.is_empty()
                    && blank_cols_before.len() as u32 > self.max_blank_column
                {
                    let ranges = format_column_ranges(&blank_cols_before);
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
        // Check if row has any non-empty data
        let has_data = (min_col..=max_col).any(|col| {
            sheet
                .cells
                .get(&(row, col))
                .is_some_and(|c| !c.value.is_empty())
        });

        // Check if row is part of a merged cell
        let in_merged_cell = sheet
            .merged_cells
            .iter()
            .any(|(r1, c1, r2, c2)| row >= *r1 && row <= *r2 && *c1 <= max_col && *c2 >= min_col);

        if !has_data && !in_merged_cell {
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
        // Check if column has any non-empty data
        let has_data = (min_row..=max_row).any(|row| {
            sheet
                .cells
                .get(&(row, col))
                .is_some_and(|c| !c.value.is_empty())
        });

        // Check if column is part of a merged cell
        let in_merged_cell = sheet
            .merged_cells
            .iter()
            .any(|(r1, c1, r2, c2)| col >= *c1 && col <= *c2 && *r1 <= max_row && *r2 >= min_row);

        if !has_data && !in_merged_cell {
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
            visible: true,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
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
            visible: true,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
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
            visible: true,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        let rule = BlankRowsColumnsRule {
            max_blank_row: 2,
            max_blank_column: 2,
        };
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_merged_cells_not_blank() {
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
            (0, 1),
            Cell {
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Text("B1".to_string()),
            },
        );
        // Row 1: blank but part of merged cell F2:F5
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
        // Merged cell F2:F5 (row 1-4, col 5) - only first cell has data
        cells.insert(
            (1, 5),
            Cell {
                num_fmt: None,
                row: 1,
                col: 5,
                value: CellValue::Text("Merged".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((5, 6)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: vec![(1, 5, 4, 5)], // F2:F5 (rows 1-4, col 5)
            sheet_path: None,
            formula_parsing_error: None,
            conditional_formatting_count: 0,
            conditional_formatting_ranges: Vec::new(),
            visible: true,
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        let rule = BlankRowsColumnsRule {
            max_blank_row: 0,
            max_blank_column: 0,
        };
        let violations = rule.check(&workbook).unwrap();

        // Rows 1, 3, 4 are blank in columns A-B range
        // But rows 1, 3, 4 are part of merged cell F2:F5, so they should NOT be reported
        // The key test: violations should not mention rows 2, 4, 5 (1-based: 3, 5, 6)
        // because they're in the merged cell range
        for violation in &violations {
            // Row 2 (1-based: 3) should NOT appear because it's in merged cell
            assert!(
                !violation.message.contains("3"),
                "Row 3 (1-based) should not be reported as blank - it's in merged cell F2:F5"
            );
        }
    }

    #[test]
    fn test_empty_sheet_no_violations() {
        let sheet = Sheet {
            name: "Empty".to_string(),
            cells: HashMap::new(),
            used_range: Some((1, 1)), // A1 reported by parser
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
            ..Default::default()
        };

        let rule = BlankRowsColumnsRule {
            max_blank_row: 2,
            max_blank_column: 2,
        };
        let violations = rule.check(&workbook).unwrap();

        // Should be skipped because cells is empty
        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_styled_but_empty_row_col() {
        let mut cells = HashMap::new();
        // Row 0: data
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text("A1".to_string()),
            },
        );
        // Row 1: Styled but Empty. Should be reported as blank row!
        cells.insert(
            (1, 0),
            Cell {
                num_fmt: Some("custom".to_string()),
                row: 1,
                col: 0,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((2, 1)), // 2 rows, 1 col
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
            ..Default::default()
        };

        let rule = BlankRowsColumnsRule {
            max_blank_row: 0,
            max_blank_column: 0,
        };
        let violations = rule.check(&workbook).unwrap();

        // Should find 1 blank row (Row 2, index 1)
        assert_eq!(violations.len(), 1);
        assert!(violations[0].message.contains("Blank rows"));
        assert!(violations[0].message.contains("2")); // Row 2 (1-based)
    }

    #[test]
    fn test_column_index_to_letter() {
        assert_eq!(column_index_to_letter(0), "A");
        assert_eq!(column_index_to_letter(25), "Z");
        assert_eq!(column_index_to_letter(26), "AA");
        assert_eq!(column_index_to_letter(27), "AB");
    }
}

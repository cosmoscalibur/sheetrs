//! SEC003: Hidden columns and rows detection

use super::{LinterRule, RuleCategory};
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;

pub struct HiddenColumnsRowsRule;

impl LinterRule for HiddenColumnsRowsRule {
    fn id(&self) -> &str {
        "SEC003"
    }

    fn name(&self) -> &str {
        "Hidden columns or rows"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::SecurityAndPrivacy
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();

        for sheet in &workbook.sheets {
            // Check hidden columns
            if !sheet.hidden_columns.is_empty() {
                let ranges = group_contiguous_indices(&sheet.hidden_columns);
                for range in ranges {
                    let range_str = format_column_range(&range);
                    violations.push(Violation::new(
                        self.id(),
                        ViolationScope::Sheet(sheet.name.clone()),
                        format!("Hidden columns: {}", range_str),
                        Severity::Warning,
                    ));
                }
            }

            // Check hidden rows
            if !sheet.hidden_rows.is_empty() {
                let ranges = group_contiguous_indices(&sheet.hidden_rows);
                for range in ranges {
                    let range_str = format_row_range(&range);
                    violations.push(Violation::new(
                        self.id(),
                        ViolationScope::Sheet(sheet.name.clone()),
                        format!("Hidden rows: {}", range_str),
                        Severity::Warning,
                    ));
                }
            }
        }

        Ok(violations)
    }
}

/// Group contiguous indices into ranges
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

/// Format column range (e.g., "A:C" or "A")
fn format_column_range(indices: &[u32]) -> String {
    if indices.is_empty() {
        return String::new();
    }

    let start_col = column_index_to_letter(indices[0]);
    if indices.len() == 1 {
        return start_col;
    }

    let end_col = column_index_to_letter(*indices.last().unwrap());
    format!("{}:{}", start_col, end_col)
}

/// Format row range (e.g., "1:3" or "1")
fn format_row_range(indices: &[u32]) -> String {
    if indices.is_empty() {
        return String::new();
    }

    let start_row = indices[0] + 1; // Convert to 1-based
    if indices.len() == 1 {
        return start_row.to_string();
    }

    let end_row = indices.last().unwrap() + 1; // Convert to 1-based
    format!("{}:{}", start_row, end_row)
}

/// Convert column index (0-based) to letter (A, B, ..., Z, AA, AB, ...)
fn column_index_to_letter(mut index: u32) -> String {
    let mut result = String::new();
    index += 1; // Convert to 1-based for calculation

    while index > 0 {
        index -= 1;
        let remainder = (index % 26) as u8;
        result.insert(0, (b'A' + remainder) as char);
        index /= 26;
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::Sheet;
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_hidden_columns() {
        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![Sheet {
                name: "Sheet1".to_string(),
                cells: HashMap::new(),
                used_range: None,
                hidden_columns: vec![0, 1, 2, 5], // A, B, C, F
                hidden_rows: Vec::new(),
                merged_cells: Vec::new(),
                sheet_path: None,
                formula_parsing_error: None,
                conditional_formatting_count: 0,
                conditional_formatting_ranges: Vec::new(),
                visible: true,
            }],
            ..Default::default()
        };

        let rule = HiddenColumnsRowsRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 2); // Two ranges: A:C and F
        assert_eq!(violations[0].rule_id, "SEC003");
        assert!(violations[0].message.contains("A:C"));
        assert!(violations[1].message.contains("F"));
    }

    #[test]
    fn test_hidden_rows() {
        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![Sheet {
                name: "Sheet1".to_string(),
                cells: HashMap::new(),
                used_range: None,
                hidden_columns: Vec::new(),
                hidden_rows: vec![0, 1, 2, 10, 11], // 1, 2, 3, 11, 12
                merged_cells: Vec::new(),
                sheet_path: None,
                formula_parsing_error: None,
                conditional_formatting_count: 0,
                conditional_formatting_ranges: Vec::new(),
                visible: true,
            }],
            ..Default::default()
        };

        let rule = HiddenColumnsRowsRule;
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 2); // Two ranges: 1:3 and 11:12
        assert_eq!(violations[0].rule_id, "SEC003");
        assert!(violations[0].message.contains("1:3"));
        assert!(violations[1].message.contains("11:12"));
    }

    #[test]
    fn test_column_index_to_letter() {
        assert_eq!(column_index_to_letter(0), "A");
        assert_eq!(column_index_to_letter(25), "Z");
        assert_eq!(column_index_to_letter(26), "AA");
        assert_eq!(column_index_to_letter(27), "AB");
    }
}

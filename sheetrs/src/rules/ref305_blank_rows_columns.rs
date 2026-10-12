//! REF305: Blank rows/columns in used ranges
//!
//! Description: Identifies large gaps of empty rows or columns within the active data area.
//! Optimized walker approach: builds `HashSet<u32>` of occupied rows/cols during cell walk,
//! then computes blanks as `used_range - occupied` in `on_sheet_end`.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::parser_utils::column_letters;
use crate::reader::{Cell, Sheet};
use crate::violation::{FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope};
use std::collections::{HashMap, HashSet};
use std::fmt::{Display, Write};

/// Threshold for contiguous blank rows to trigger a violation.
const MAX_BLANK_ROW: u32 = 2;
/// Threshold for contiguous blank columns to trigger a violation.
const MAX_BLANK_COLUMN: u32 = 2;

/// Per-sheet occupied tracking.
struct SheetOccupancy {
    /// Rows that contain at least one non-empty cell.
    occupied_rows: HashSet<u32>,
    /// Columns that contain at least one non-empty cell.
    occupied_cols: HashSet<u32>,
    /// Minimum row seen.
    min_row: u32,
    /// Minimum column seen.
    min_col: u32,
}

impl Default for SheetOccupancy {
    fn default() -> Self {
        Self {
            occupied_rows: HashSet::new(),
            occupied_cols: HashSet::new(),
            min_row: u32::MAX,
            min_col: u32::MAX,
        }
    }
}

impl SheetOccupancy {
    fn record(&mut self, cell: &Cell) {
        self.occupied_rows.insert(cell.row);
        self.occupied_cols.insert(cell.col);
        self.min_row = self.min_row.min(cell.row);
        self.min_col = self.min_col.min(cell.col);
    }
}

/// Rule that detects blank rows and columns within used ranges.
///
/// Uses a constant threshold of 2 to identify gaps of empty rows or columns.
/// Optimized: tracks occupied rows/cols during cell walk instead of probing
/// every (row, col) pair.
pub struct BlankRowsColumnsRule {
    /// Per-sheet occupancy data.
    sheet_data: HashMap<u16, SheetOccupancy>,
}

impl BlankRowsColumnsRule {
    /// Create a new instance.
    pub fn new() -> Self {
        Self {
            sheet_data: HashMap::new(),
        }
    }
}

impl Default for BlankRowsColumnsRule {
    fn default() -> Self {
        Self::new()
    }
}

/// Incident data for REF305: blank rows.
#[derive(Debug)]
pub struct BlankRowsData {
    /// Formatted range string (e.g., "2-4, 8").
    pub ranges: String,
}

impl ViolationData for BlankRowsData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!(
            "Blank rows within used range: {}. Consider removing or filling these rows.",
            self.ranges
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Incident data for REF305: blank columns.
#[derive(Debug)]
pub struct BlankColumnsData {
    /// Formatted range string (e.g., "B-D, H").
    pub ranges: String,
}

impl ViolationData for BlankColumnsData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        format!(
            "Blank columns within used range: {}. Consider removing or filling these columns.",
            self.ranges
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for BlankRowsColumnsRule {
    fn id(&self) -> RuleId {
        RuleId::Ref305
    }

    fn name(&self) -> &'static str {
        "Blank Row or Column"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Reference
    }

    fn on_cell(&mut self, sheet: &Sheet, cell: &Cell, _ctx: &mut LinterContext) -> Vec<Violation> {
        if !cell.value.is_empty() {
            self.sheet_data
                .entry(sheet.sheet_index)
                .or_default()
                .record(cell);
        }
        Vec::new()
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext) -> Vec<Violation> {
        let Some(data) = self.sheet_data.remove(&sheet.sheet_index) else {
            return Vec::new(); // No non-empty cells → skip
        };

        let (max_row, max_col) = match sheet.used_range {
            Some((ur, uc)) => (ur.saturating_sub(1), uc.saturating_sub(1)),
            None => return Vec::new(),
        };

        let min_row = data.min_row;
        let min_col = data.min_col;

        // Build merged-cell coverage sets for exclusion
        let mut merged_rows: HashSet<u32> = HashSet::new();
        let mut merged_cols: HashSet<u32> = HashSet::new();
        for &(r1, c1, r2, c2) in &sheet.merged_cells {
            for r in r1..=r2 {
                if c1 <= max_col && c2 >= min_col {
                    merged_rows.insert(r);
                }
            }
            for c in c1..=c2 {
                if r1 <= max_row && r2 >= min_row {
                    merged_cols.insert(c);
                }
            }
        }

        let mut violations = Vec::new();

        // Blank rows BEFORE data starts (from row 0)
        if min_row > MAX_BLANK_ROW {
            let blank_before: Vec<u32> = (0..min_row).collect();
            violations.push(Violation::with_data(
                RuleId::Ref305,
                ViolationScope::Sheet(sheet.sheet_index),
                BlankRowsData {
                    ranges: format_row_ranges(&blank_before),
                },
                Severity::Info,
            ));
        }

        // Blank columns BEFORE data starts (from col 0)
        if min_col > MAX_BLANK_COLUMN {
            let blank_before: Vec<u32> = (0..min_col).collect();
            violations.push(Violation::with_data(
                RuleId::Ref305,
                ViolationScope::Sheet(sheet.sheet_index),
                BlankColumnsData {
                    ranges: format_column_ranges(&blank_before),
                },
                Severity::Info,
            ));
        }

        // Blank rows WITHIN used range
        let blank_rows: Vec<u32> = (min_row..=max_row)
            .filter(|r| !data.occupied_rows.contains(r) && !merged_rows.contains(r))
            .collect();

        if !blank_rows.is_empty() {
            let groups = group_contiguous_indices(&blank_rows);
            let filtered: Vec<u32> = groups
                .into_iter()
                .filter(|g| g.len() as u32 > MAX_BLANK_ROW)
                .flatten()
                .collect();

            if !filtered.is_empty() {
                violations.push(Violation::with_data(
                    RuleId::Ref305,
                    ViolationScope::Sheet(sheet.sheet_index),
                    BlankRowsData {
                        ranges: format_row_ranges(&filtered),
                    },
                    Severity::Info,
                ));
            }
        }

        // Blank columns WITHIN used range
        let blank_cols: Vec<u32> = (min_col..=max_col)
            .filter(|c| !data.occupied_cols.contains(c) && !merged_cols.contains(c))
            .collect();

        if !blank_cols.is_empty() {
            let groups = group_contiguous_indices(&blank_cols);
            let filtered: Vec<u32> = groups
                .into_iter()
                .filter(|g| g.len() as u32 > MAX_BLANK_COLUMN)
                .flatten()
                .collect();

            if !filtered.is_empty() {
                violations.push(Violation::with_data(
                    RuleId::Ref305,
                    ViolationScope::Sheet(sheet.sheet_index),
                    BlankColumnsData {
                        ranges: format_column_ranges(&filtered),
                    },
                    Severity::Info,
                ));
            }
        }

        violations
    }
}

/// Group contiguous indices into ranges.
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

/// Format row ranges (e.g., "2, 4-6, 10").
fn format_row_ranges(rows: &[u32]) -> String {
    format_ranges(rows, |r| r + 1) // Convert to 1-based
}

/// Format column ranges (e.g., "C, E-G, J").
fn format_column_ranges(cols: &[u32]) -> String {
    format_ranges(cols, column_letters)
}

/// Generic range formatter.
fn format_ranges<D, F>(indices: &[u32], formatter: F) -> String
where
    D: Display,
    F: Fn(u32) -> D,
{
    let mut out = String::new();
    let Some((&first, rest)) = indices.split_first() else {
        return out;
    };

    let mut push_range = |start: u32, end: u32| {
        if !out.is_empty() {
            out.push_str(", ");
        }
        let _ = if start == end {
            write!(out, "{}", formatter(start))
        } else {
            write!(out, "{}-{}", formatter(start), formatter(end))
        };
    };

    let mut start = first;
    let mut end = first;
    for &idx in rest {
        if idx == end + 1 {
            end = idx;
        } else {
            push_range(start, end);
            start = idx;
            end = idx;
        }
    }
    push_range(start, end);

    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;
    use std::sync::Arc;

    #[test]
    fn test_blank_rows() {
        let mut cells = HashMap::new();
        // Row 0: A1, B1
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text(Arc::from("A1")),
            },
        );
        cells.insert(
            (0, 1),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Text(Arc::from("B1")),
            },
        );
        // Rows 1, 2, 3: blank → 3 contiguous (exceeds threshold of 2)
        // Row 4: A5, B5
        cells.insert(
            (4, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 4,
                col: 0,
                value: CellValue::Text(Arc::from("A5")),
            },
        );
        cells.insert(
            (4, 1),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 4,
                col: 1,
                value: CellValue::Text(Arc::from("B5")),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((5, 2)),
            ..Default::default()
        };

        let mut rule = BlankRowsColumnsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        // With threshold of 2, should catch 3 blank rows
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref305);
        assert!(violations[0].message().contains("Blank rows"));
        // Should contain rows 2, 3, 4 (1-based)
        assert!(violations[0].message().contains("2-4"));
    }

    #[test]
    fn test_blank_columns() {
        let mut cells = HashMap::new();
        // Column A: A1, A2
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text(Arc::from("A1")),
            },
        );
        cells.insert(
            (1, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 1,
                col: 0,
                value: CellValue::Text(Arc::from("A2")),
            },
        );
        // Columns B, C, D: blank → 3 contiguous (exceeds threshold of 2)
        // Column E: E1, E2
        cells.insert(
            (0, 4),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 4,
                value: CellValue::Text(Arc::from("E1")),
            },
        );
        cells.insert(
            (1, 4),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 1,
                col: 4,
                value: CellValue::Text(Arc::from("E2")),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((2, 5)),
            ..Default::default()
        };

        let mut rule = BlankRowsColumnsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        // With threshold of 2, should catch 3 blank columns
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref305);
        assert!(violations[0].message().contains("Blank columns"));
        // Should contain columns B, C, D
        assert!(violations[0].message().contains("B-D"));
    }

    #[test]
    fn test_no_blanks() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text(Arc::from("A1")),
            },
        );
        cells.insert(
            (0, 1),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Text(Arc::from("B1")),
            },
        );
        cells.insert(
            (1, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 1,
                col: 0,
                value: CellValue::Text(Arc::from("A2")),
            },
        );
        cells.insert(
            (1, 1),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 1,
                col: 1,
                value: CellValue::Text(Arc::from("B2")),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((2, 2)),
            ..Default::default()
        };

        let mut rule = BlankRowsColumnsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert!(violations.is_empty());
    }

    #[test]
    fn test_one_blank_row_no_violation() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text(Arc::from("A1")),
            },
        );
        cells.insert(
            (0, 1),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Text(Arc::from("B1")),
            },
        );
        // Row 1: blank (only 1)
        cells.insert(
            (2, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 2,
                col: 0,
                value: CellValue::Text(Arc::from("A3")),
            },
        );
        cells.insert(
            (2, 1),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 2,
                col: 1,
                value: CellValue::Text(Arc::from("B3")),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((3, 2)),
            ..Default::default()
        };

        let mut rule = BlankRowsColumnsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert!(violations.is_empty());
    }

    #[test]
    fn test_two_blank_rows_no_violation() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text(Arc::from("A1")),
            },
        );
        cells.insert(
            (0, 1),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::Text(Arc::from("B1")),
            },
        );
        // Rows 1, 2: blank (exactly 2, at threshold)
        cells.insert(
            (3, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 3,
                col: 0,
                value: CellValue::Text(Arc::from("A4")),
            },
        );
        cells.insert(
            (3, 1),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 3,
                col: 1,
                value: CellValue::Text(Arc::from("B4")),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((4, 2)),
            ..Default::default()
        };

        let mut rule = BlankRowsColumnsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        // Threshold is >, not >=, so exactly 2 should NOT trigger
        assert!(violations.is_empty());
    }

    #[test]
    fn test_one_blank_column_no_violation() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text(Arc::from("A1")),
            },
        );
        cells.insert(
            (1, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 1,
                col: 0,
                value: CellValue::Text(Arc::from("A2")),
            },
        );
        // Column B: blank (only 1)
        cells.insert(
            (0, 2),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 2,
                value: CellValue::Text(Arc::from("C1")),
            },
        );
        cells.insert(
            (1, 2),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 1,
                col: 2,
                value: CellValue::Text(Arc::from("C2")),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((2, 3)),
            ..Default::default()
        };

        let mut rule = BlankRowsColumnsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert!(violations.is_empty());
    }

    #[test]
    fn test_two_blank_columns_no_violation() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text(Arc::from("A1")),
            },
        );
        cells.insert(
            (1, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 1,
                col: 0,
                value: CellValue::Text(Arc::from("A2")),
            },
        );
        // Columns B, C: blank (exactly 2, at threshold)
        cells.insert(
            (0, 3),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 3,
                value: CellValue::Text(Arc::from("D1")),
            },
        );
        cells.insert(
            (1, 3),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 1,
                col: 3,
                value: CellValue::Text(Arc::from("D2")),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((2, 4)),
            ..Default::default()
        };

        let mut rule = BlankRowsColumnsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert!(violations.is_empty());
    }

    #[test]
    fn test_merged_cells_not_blank() {
        let mut cells = HashMap::new();
        // Row 0: A1
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text(Arc::from("A1")),
            },
        );
        // Rows 1-3: no data, but merged cell covers B2:B4
        // Row 4: A5
        cells.insert(
            (4, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 4,
                col: 0,
                value: CellValue::Text(Arc::from("A5")),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            used_range: Some((5, 2)),
            merged_cells: vec![(1, 1, 3, 1)], // B2:B4 merged
            ..Default::default()
        };

        let mut rule = BlankRowsColumnsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        // Rows 1-3 are covered by merged cell, should NOT be blank
        assert!(violations.is_empty());
    }

    #[test]
    fn test_empty_sheet_no_violations() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: HashMap::new(),
            used_range: Some((100, 100)),
            ..Default::default()
        };

        let mut rule = BlankRowsColumnsRule::new();
        let mut ctx = LinterContext::default();
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert!(violations.is_empty());
    }
}

//! REF307: Avoid whole-column or whole-row references
//!
//! Description: Detects whole-column (e.g., A:A) or whole-row (e.g., 1:1) references.

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use regex::Regex;
use std::collections::HashMap;
use std::sync::{LazyLock, Mutex};

/// Matches whole-column references (A:A, A:Z, etc.).
static COLUMN_PATTERN: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"\b[A-Z]+:[A-Z]+\b").expect("REF307 column regex must compile"));

/// Matches whole-row references (1:1, 1:100, etc.).
static ROW_PATTERN: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"\b\d+:\d+\b").expect("REF307 row regex must compile"));

/// Per-sheet cell collection: sheet_index → Vec<(row, col, is_column)>.
type SheetCellMap = Mutex<HashMap<u16, Vec<(u32, u32, bool)>>>;

/// Rule that detects whole-column (e.g., A:A) or whole-row (e.g., 1:1) references.
pub struct WholeColumnRowRefsRule {
    /// Per-sheet collected cells.
    sheet_cells: SheetCellMap,
}

impl WholeColumnRowRefsRule {
    /// Create a new instance.
    pub fn new() -> Self {
        Self {
            sheet_cells: Mutex::new(HashMap::new()),
        }
    }
}

impl Default for WholeColumnRowRefsRule {
    fn default() -> Self {
        Self::new()
    }
}

/// Incident data for REF307.
#[derive(Debug)]
pub struct WholeRefData {
    /// `true` for whole-column, `false` for whole-row.
    pub is_column: bool,
    /// Bounding box of violating cells (start_row, start_col, end_row, end_col).
    pub range: (u32, u32, u32, u32),
}

impl ViolationData for WholeRefData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let (sr, sc, er, ec) = self.range;
        let range_str = if sr == er && sc == ec {
            CellReference::new(sr, sc).to_string()
        } else {
            format!(
                "{}:{}",
                CellReference::new(sr, sc),
                CellReference::new(er, ec)
            )
        };

        if self.is_column {
            format!(
                "Whole-column reference (e.g., A:A) found in range: {range_str}. Use bounded ranges for better performance."
            )
        } else {
            format!(
                "Whole-row reference (e.g., 1:1) found in range: {range_str}. Use bounded ranges for better performance."
            )
        }
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for WholeColumnRowRefsRule {
    fn id(&self) -> RuleId {
        RuleId::Ref307
    }

    fn name(&self) -> &'static str {
        "Whole Column or Row Reference"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Reference
    }

    fn on_cell(&self, sheet: &Sheet, cell: &Cell, ctx: &mut LinterContext) -> Vec<Violation> {
        if let Some(formula_upper) = ctx.formula_upper(sheet, cell) {
            let has_column = COLUMN_PATTERN.is_match(formula_upper);
            let has_row = !has_column && ROW_PATTERN.is_match(formula_upper);

            if has_column || has_row {
                let mut map = self.sheet_cells.lock().unwrap();
                map.entry(sheet.sheet_index)
                    .or_default()
                    .push((cell.row, cell.col, has_column));
            }
        }
        Vec::new()
    }

    fn on_sheet_end(&self, sheet: &Sheet, _ctx: &mut LinterContext) -> Vec<Violation> {
        let Some(cells) = self.sheet_cells.lock().unwrap().remove(&sheet.sheet_index) else {
            return Vec::new();
        };

        let mut violations = Vec::new();

        // Split by type
        let column_cells: Vec<(u32, u32)> =
            cells.iter().filter(|c| c.2).map(|c| (c.0, c.1)).collect();
        let row_cells: Vec<(u32, u32)> =
            cells.iter().filter(|c| !c.2).map(|c| (c.0, c.1)).collect();

        for (is_column, group) in [(true, &column_cells), (false, &row_cells)] {
            if group.is_empty() {
                continue;
            }
            let ranges = find_contiguous_ranges(group);
            for range in ranges {
                let bbox = bounding_box(&range);
                violations.push(Violation::with_data(
                    RuleId::Ref307,
                    ViolationScope::Sheet(sheet.sheet_index),
                    WholeRefData {
                        is_column,
                        range: bbox,
                    },
                    Severity::Warning,
                ));
            }
        }

        violations
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;

    #[test]
    fn test_whole_column_reference() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=SUM(A:A)")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            ..Default::default()
        };

        let rule = WholeColumnRowRefsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref307);
    }

    #[test]
    fn test_whole_row_reference() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=SUM(1:1)")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            ..Default::default()
        };

        let rule = WholeColumnRowRefsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref307);
    }

    #[test]
    fn test_bounded_reference() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=SUM(A1:A10)")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            ..Default::default()
        };

        let rule = WholeColumnRowRefsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert!(violations.is_empty());
    }
}

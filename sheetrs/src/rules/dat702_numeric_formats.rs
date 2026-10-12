//! DATA702: Inconsistent number formatting detection
//!
//! Description: Detects numeric data stored as text instead of as number type.

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, CellValue, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::HashMap;

/// Per-sheet cell collection: sheet_index → Vec<(row, col)>.
type SheetCellMap = HashMap<u16, Vec<(u32, u32)>>;

/// Rule that detects numbers stored as text
pub struct InconsistentNumberFormatRule {
    /// Per-sheet collected cells.
    sheet_cells: SheetCellMap,
}

impl InconsistentNumberFormatRule {
    /// Create a new instance.
    pub fn new() -> Self {
        Self {
            sheet_cells: HashMap::new(),
        }
    }
}

impl Default for InconsistentNumberFormatRule {
    fn default() -> Self {
        Self::new()
    }
}

/// Incident data for DATA702.
#[derive(Debug)]
pub struct NumericTextData {
    /// Bounding box of violating cells (start_row, start_col, end_row, end_col).
    pub range: (u32, u32, u32, u32),
}

impl ViolationData for NumericTextData {
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
        format!("Numeric data stored as text in range: {range_str}")
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Check if a text string represents a numeric value.
fn is_numeric_text(text: &str) -> bool {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return false;
    }
    trimmed.parse::<f64>().is_ok()
}

impl WalkerRule for InconsistentNumberFormatRule {
    fn id(&self) -> RuleId {
        RuleId::Data702
    }

    fn name(&self) -> &'static str {
        "Number Stored as Text"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Data
    }

    fn on_cell(&mut self, sheet: &Sheet, cell: &Cell, _ctx: &mut LinterContext) -> Vec<Violation> {
        if let CellValue::Text(text) = &cell.value
            && is_numeric_text(text)
        {
            self.sheet_cells
                .entry(sheet.sheet_index)
                .or_default()
                .push((cell.row, cell.col));
        }
        Vec::new()
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext) -> Vec<Violation> {
        let Some(mut cells) = self.sheet_cells.remove(&sheet.sheet_index) else {
            return Vec::new();
        };
        cells.sort_unstable();

        let ranges = find_contiguous_ranges(&cells);
        let mut violations = Vec::new();

        for range in ranges {
            let bbox = bounding_box(&range);
            violations.push(Violation::with_data(
                RuleId::Data702,
                ViolationScope::Sheet(sheet.sheet_index),
                NumericTextData { range: bbox },
                Severity::Warning,
            ));
        }

        violations
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;
    use std::sync::Arc;

    #[test]
    fn test_numeric_text_detection() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text(Arc::from("42")),
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
                value: CellValue::Text(Arc::from("3.14")),
            },
        );
        // Real number — not flagged
        cells.insert(
            (2, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 2,
                col: 0,
                value: CellValue::Number(100.0),
            },
        );
        // Non-numeric text — not flagged
        cells.insert(
            (3, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                row: 3,
                col: 0,
                value: CellValue::Text(Arc::from("Hello")),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            ..Default::default()
        };

        let mut rule = InconsistentNumberFormatRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Data702);
    }
}

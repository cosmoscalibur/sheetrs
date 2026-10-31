//! DATA702: Number stored as text
//!
//! A number kept as text (`'123`, `"007"` typed into a Text-formatted cell) is ignored by `SUM`,
//! sorts and compares as text and breaks lookups against real numbers; the cell is usually the
//! result of an import or a format applied before typing.
//!
//! # Reports
//!
//! Constant cells whose value is text that parses as a number (after trimming spaces). Adjacent
//! cells of the same sheet are merged into one violation per contiguous range, reported with the
//! bounding box of the range (`D21:BI21`).
//!
//! # Does not report
//!
//! - **Text a formula returns** (`=IF(A1,"0","1")`, `=CONCAT(...)` of digits, `=A1` of a text
//!   cell): that is a result and not stored data, and the cell cannot be converted without
//!   changing the formula. The stored cell it comes from is the one reported, and
//!   [DATA706](super::dat706_numeric_text_calc) reports the formulas that calculate with it.
//! - **A number in a cell formatted as Text** (`<v>210</v>` with format `@`): the stored value is
//!   a number whatever the format says.
//! - Text that is not numeric (`"A1"`, `"1,5 kg"`) and empty or blank text.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Warning.

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

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        at: (u32, u32),
        cell: &Cell,
        _ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        if let CellValue::Text(text) = &cell.value
            && !cell.is_formula()
            && is_numeric_text(text)
        {
            self.sheet_cells
                .entry(sheet.sheet_index)
                .or_default()
                .push((at.0, at.1));
        }
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext, out: &mut Vec<Violation>) {
        let Some(mut cells) = self.sheet_cells.remove(&sheet.sheet_index) else {
            return;
        };
        cells.sort_unstable();

        let ranges = find_contiguous_ranges(&cells);

        for range in ranges {
            let bbox = bounding_box(&range);
            out.push(Violation::with_data(
                RuleId::Data702,
                ViolationScope::Sheet(sheet.sheet_index),
                NumericTextData { range: bbox },
                Severity::Warning,
            ));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::collect_violations;
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
                value: CellValue::Text(Arc::from("42")),
            },
        );
        cells.insert(
            (1, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
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
                value: CellValue::Text(Arc::from("Hello")),
            },
        );

        // Numeric text that a formula returns is a result, not stored data: not flagged.
        cells.insert(
            (4, 0),
            Cell {
                formula: Some("IF(A1,\"0\",\"1\")".into()),
                is_array: false,
                num_fmt: None,
                value: CellValue::Text(Arc::from("0")),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            ..Default::default()
        };

        let mut rule = InconsistentNumberFormatRule::new();
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Data702);
    }
}

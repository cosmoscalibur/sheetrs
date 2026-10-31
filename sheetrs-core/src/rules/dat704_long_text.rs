//! DATA704: Long text cells
//!
//! Very long text in a cell is usually data put in the wrong place (logs, JSON, pasted documents)
//! or text that should be split or stored elsewhere. Such cells are unreadable on screen, slow to
//! load and hard to maintain.
//!
//! # Reports
//!
//! One violation per range (unit: a range of cells, reported at sheet level). Every text cell whose
//! length is greater than the threshold is collected; cells that touch horizontally or vertically
//! (4-connected) form one group, and each group is reported once as its bounding box, such as
//! `B2:B9` or a single cell. Length is the number of characters (Unicode scalar values) of the text, so an accented letter or an
//! emoji counts as one, however many bytes it takes. A group of an irregular shape is reported by its bounding box, which can
//! contain cells that are not long. Text that a formula returns counts like any other text value:
//! the code does not distinguish it.
//!
//! # Does not report
//!
//! - Text of exactly the threshold length or shorter.
//! - Numbers, dates, booleans and errors: only text values are measured.
//! - Long formulas: that is CPX509, which measures the formula and not its result.
//!
//! # Parameters
//!
//! - `max_text_length` (integer, default `255`): the largest accepted length in characters. It can be set
//!   per sheet.
//!
//! # Severity
//!
//! Always `Warning`.

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::{Cell, CellValue, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::HashMap;

/// Per-sheet cell collection: sheet_index → Vec<(row, col)>.
type SheetCellMap = HashMap<u16, Vec<(u32, u32)>>;

/// Rule that detects cells with excessive text length
pub struct LongTextCellRule {
    /// Rule configuration.
    config: LinterConfig,
    /// Per-sheet collected cells.
    sheet_cells: SheetCellMap,
}

impl LongTextCellRule {
    /// Create a new instance with configuration.
    pub fn new(config: &LinterConfig) -> Self {
        Self {
            config: config.clone(),
            sheet_cells: HashMap::new(),
        }
    }
}

/// Incident data for DATA704.
#[derive(Debug)]
pub struct LongTextData {
    /// Bounding box of violating cells (start_row, start_col, end_row, end_col).
    pub range: (u32, u32, u32, u32),
}

impl ViolationData for LongTextData {
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
        format!("Long text cells in range: {range_str}")
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for LongTextCellRule {
    fn id(&self) -> RuleId {
        RuleId::Data704
    }

    fn name(&self) -> &'static str {
        "Long text cell"
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
        let threshold = self
            .config
            .get_param_int("max_text_length", Some(&sheet.name))
            .unwrap_or(255) as usize;

        if let CellValue::Text(text) = &cell.value
            && text.chars().count() > threshold
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
                RuleId::Data704,
                ViolationScope::Sheet(sheet.sheet_index),
                LongTextData { range: bbox },
                Severity::Warning,
            ));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::Sheet;
    use crate::rules::collect_violations;
    use std::collections::HashMap;
    use std::sync::Arc;

    #[test]
    fn test_long_text_cell() {
        let mut cells = HashMap::new();
        let long_text = "A".repeat(300); // >255 chars

        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                value: CellValue::Text(Arc::from(long_text.as_str())),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            ..Default::default()
        };

        let mut rule = LongTextCellRule::new(&LinterConfig::default());
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Data704);
    }

    #[test]
    fn length_is_in_characters_not_bytes() {
        // 100 characters of 3 bytes each: 300 bytes, under the 255-character limit.
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                value: CellValue::Text(Arc::from("日".repeat(100).as_str())),
            },
        );
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            ..Default::default()
        };
        let mut rule = LongTextCellRule::new(&LinterConfig::default());
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        assert!(collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out)).is_empty());
    }
}

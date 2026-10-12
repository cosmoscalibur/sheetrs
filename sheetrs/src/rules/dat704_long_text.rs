//! DATA704: Long text cell detection
//!
//! Description: Extremely long text strings in cells are often data misuse
//! (storing logs/JSON in cells) or formatted incorrectly.

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::{Cell, CellValue, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::HashMap;
use std::sync::Mutex;

/// Per-sheet cell collection: sheet_index → Vec<(row, col)>.
type SheetCellMap = Mutex<HashMap<u16, Vec<(u32, u32)>>>;

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
            sheet_cells: Mutex::new(HashMap::new()),
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

    fn on_cell(&self, sheet: &Sheet, cell: &Cell, _ctx: &mut LinterContext) -> Vec<Violation> {
        let threshold = self
            .config
            .get_param_int("max_text_length", Some(&sheet.name))
            .unwrap_or(255) as usize;

        if let CellValue::Text(text) = &cell.value
            && text.len() > threshold
        {
            let mut map = self.sheet_cells.lock().unwrap();
            map.entry(sheet.sheet_index)
                .or_default()
                .push((cell.row, cell.col));
        }
        Vec::new()
    }

    fn on_sheet_end(&self, sheet: &Sheet, _ctx: &mut LinterContext) -> Vec<Violation> {
        let Some(cells) = self.sheet_cells.lock().unwrap().remove(&sheet.sheet_index) else {
            return Vec::new();
        };

        let ranges = find_contiguous_ranges(&cells);
        let mut violations = Vec::new();

        for range in ranges {
            let bbox = bounding_box(&range);
            violations.push(Violation::with_data(
                RuleId::Data704,
                ViolationScope::Sheet(sheet.sheet_index),
                LongTextData { range: bbox },
                Severity::Warning,
            ));
        }

        violations
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::Sheet;
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
                row: 0,
                col: 0,
                value: CellValue::Text(Arc::from(long_text.as_str())),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells,
            ..Default::default()
        };

        let rule = LongTextCellRule::new(&LinterConfig::default());
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Data704);
    }
}

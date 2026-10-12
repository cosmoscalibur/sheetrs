//! EXT801: External workbook references
//!
//! Description: Identifies cell-level formula links to external spreadsheet files.

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use regex::Regex;
use std::collections::{BTreeMap, HashMap};
use std::sync::LazyLock;

/// Per-sheet cell collection: sheet_index → Vec<(row, col, wb_index)>.
type SheetCellMap = HashMap<u16, Vec<(u32, u32, usize)>>;

/// Rule that detects references to external workbooks
pub struct ExternalWorkbooksRule {
    /// Cells referencing external workbooks per sheet.
    sheet_cells: SheetCellMap,
}

impl ExternalWorkbooksRule {
    /// Create a new instance of the rule.
    pub fn new() -> Self {
        Self {
            sheet_cells: HashMap::new(),
        }
    }
}

impl Default for ExternalWorkbooksRule {
    fn default() -> Self {
        Self::new()
    }
}

/// Incident data for EXT801.
#[derive(Debug)]
pub struct ExternalWorkbookData {
    /// 0-based external workbook index (into `Workbook::external_workbooks`).
    pub workbook_index: usize,
    /// Range as (start_row, start_col, end_row, end_col).
    pub range: (u32, u32, u32, u32),
}

impl ViolationData for ExternalWorkbookData {
    fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        let wb_name = ctx
            .workbook
            .external_workbooks
            .get(self.workbook_index)
            .map(|wb| wb.path.as_str())
            .unwrap_or("unknown");

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

        format!("External workbook reference {wb_name} found in range: {range_str}")
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

static INDEX_PATTERN: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"\[(\d+)\]").unwrap());

/// Extract `[N]` indices from formula
fn extract_external_workbook_indices(formula: &str) -> Vec<usize> {
    let mut indices = Vec::new();
    for cap in INDEX_PATTERN.captures_iter(formula) {
        if let Some(num_str) = cap.get(1)
            && let Ok(num) = num_str.as_str().parse::<usize>()
            && num > 0
        {
            indices.push(num - 1); // Convert 1-based to 0-based
        }
    }
    indices
}

impl WalkerRule for ExternalWorkbooksRule {
    fn id(&self) -> RuleId {
        RuleId::Ext801
    }

    fn name(&self) -> &'static str {
        "External Workbook Reference"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::External
    }

    fn on_cell(&mut self, sheet: &Sheet, cell: &Cell, _ctx: &mut LinterContext) -> Vec<Violation> {
        if let Some(formula) = cell.as_formula() {
            let indices = extract_external_workbook_indices(formula);
            if !indices.is_empty() {
                self.sheet_cells
                    .entry(sheet.sheet_index)
                    .or_default()
                    .extend(indices.into_iter().map(|idx| (cell.row, cell.col, idx)));
            }
        }
        Vec::new()
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext) -> Vec<Violation> {
        let Some(mut cells) = self.sheet_cells.remove(&sheet.sheet_index) else {
            return Vec::new();
        };
        cells.sort_unstable();

        let mut violations = Vec::new();
        // Group cells by workbook index
        let mut grouped: BTreeMap<usize, Vec<(u32, u32)>> = BTreeMap::new();
        for (row, col, idx) in cells {
            grouped.entry(idx).or_default().push((row, col));
        }

        for (idx, group_cells) in grouped {
            let ranges = find_contiguous_ranges(&group_cells);
            for range in ranges {
                violations.push(Violation::with_data(
                    RuleId::Ext801,
                    ViolationScope::Sheet(sheet.sheet_index),
                    ExternalWorkbookData {
                        workbook_index: idx,
                        range: bounding_box(&range),
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
    use crate::rules::LinterContext;
    use std::collections::HashMap;

    #[test]
    fn test_external_workbook_in_formula() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=[1]Sheet1!A1")),
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
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let mut rule = ExternalWorkbooksRule::new();
        let mut ctx = LinterContext::default();

        // Walk cells
        for cell in sheet.cells.values() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }

        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ext801);
    }
}

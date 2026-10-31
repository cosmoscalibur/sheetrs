//! EXT801: External Workbook
//!
//! A formula that reads another workbook (`=[1]Prices!B4`) depends on a file that is not part of
//! this one: it breaks when the file is moved or renamed, shows stale values when links are not
//! refreshed, and can leak a path from the author's machine.
//!
//! # Reports
//!
//! One violation per contiguous block of formula cells that reference the same external workbook
//! (unit: a range, scoped to the sheet; one-cell blocks are reported as a single cell). Cells are
//! grouped per sheet and per external workbook, and cells that touch by row or column adjacency
//! form one block, so a filled-down column of links is one violation. The message gives the path of
//! the workbook, as stored in the file, and the bounding range of the block. A formula that links to
//! several workbooks appears in each workbook's blocks.
//!
//! A reference is recognised textually by an index in square brackets, `[N]`, anywhere in the
//! formula text; `N` is the 1-based position in the list of external workbooks, and an index not in
//! that list is shown as `unknown`. `[0]` is ignored.
//!
//! # Does not report
//!
//! - Links that are not in a cell formula (defined names, charts, data validation, conditional
//!   formats).
//! - Web addresses in cell text, which are EXT802.
//! - Whether the linked file exists or has been updated; the file is never opened.
//! - Formulas inside string constants are not excluded, so text such as `"[1]"` inside a formula is
//!   read as an index.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always `Warning`.

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

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        at: (u32, u32),
        cell: &Cell,
        _ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        if let Some(formula) = cell.as_formula() {
            let indices = extract_external_workbook_indices(formula);
            if !indices.is_empty() {
                self.sheet_cells
                    .entry(sheet.sheet_index)
                    .or_default()
                    .extend(indices.into_iter().map(|idx| (at.0, at.1, idx)));
            }
        }
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext, out: &mut Vec<Violation>) {
        let Some(mut cells) = self.sheet_cells.remove(&sheet.sheet_index) else {
            return;
        };
        cells.sort_unstable();
        // Group cells by workbook index
        let mut grouped: BTreeMap<usize, Vec<(u32, u32)>> = BTreeMap::new();
        for (row, col, idx) in cells {
            grouped.entry(idx).or_default().push((row, col));
        }

        for (idx, group_cells) in grouped {
            let ranges = find_contiguous_ranges(&group_cells);
            for range in ranges {
                out.push(Violation::with_data(
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
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::LinterContext;
    use crate::rules::collect_violations;
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
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let mut rule = ExternalWorkbooksRule::new();
        let mut ctx = LinterContext::default();

        // Walk cells
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }

        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ext801);
    }
}

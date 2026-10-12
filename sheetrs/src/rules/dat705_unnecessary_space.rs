//! DATA705: Unnecessary Space detection
//!
//! Description: Detects invisible white-space padding at the start or end of cell values.

use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};

/// Incident data for DATA705.
#[derive(Debug)]
pub struct UnnecessarySpaceData;

impl ViolationData for UnnecessarySpaceData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        "Cell contains unnecessary whitespace".to_string()
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Rule that identifies unnecessary leading or trailing whitespace in text cells.
pub struct UnnecessarySpaceRule;

impl WalkerRule for UnnecessarySpaceRule {
    fn id(&self) -> RuleId {
        RuleId::Data705
    }

    fn name(&self) -> &'static str {
        "Unnecessary Space"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Data
    }

    fn on_cell(&self, sheet: &Sheet, cell: &Cell, _ctx: &mut LinterContext) -> Vec<Violation> {
        if let crate::reader::CellValue::Text(text) = &cell.value {
            let bytes = text.as_bytes();
            let has_whitespace = bytes.first().is_some_and(|b| b.is_ascii_whitespace())
                || bytes.last().is_some_and(|b| b.is_ascii_whitespace());

            if has_whitespace {
                return vec![Violation::with_data(
                    RuleId::Data705,
                    ViolationScope::Cell(sheet.sheet_index, CellReference::new(cell.row, cell.col)),
                    UnnecessarySpaceData,
                    Severity::Info,
                )];
            }
        }
        Vec::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::LinterContext;
    use std::collections::HashMap;
    use std::sync::Arc;

    fn make_text_cell(row: u32, col: u32, text: &str) -> Cell {
        Cell {
            row,
            col,
            value: CellValue::Text(Arc::from(text)),
            formula: None,
            is_array: false,
            num_fmt: None,
        }
    }

    fn make_sheet(cells: Vec<Cell>) -> Sheet {
        let mut cell_map = HashMap::new();
        for c in cells {
            cell_map.insert((c.row, c.col), c);
        }
        Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cell_map,
            ..Default::default()
        }
    }

    #[test]
    fn test_leading_space() {
        let sheet = make_sheet(vec![make_text_cell(0, 0, " hello")]);
        let rule = UnnecessarySpaceRule;
        let mut ctx = LinterContext::default();

        let violations: Vec<_> = sheet
            .all_cells()
            .flat_map(|cell| rule.on_cell(&sheet, cell, &mut ctx))
            .collect();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Data705);
    }

    #[test]
    fn test_trailing_space() {
        let sheet = make_sheet(vec![make_text_cell(0, 0, "hello ")]);
        let rule = UnnecessarySpaceRule;
        let mut ctx = LinterContext::default();

        let violations: Vec<_> = sheet
            .all_cells()
            .flat_map(|cell| rule.on_cell(&sheet, cell, &mut ctx))
            .collect();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Data705);
    }

    #[test]
    fn test_both_spaces() {
        let sheet = make_sheet(vec![make_text_cell(0, 0, " hello ")]);
        let rule = UnnecessarySpaceRule;
        let mut ctx = LinterContext::default();

        let violations: Vec<_> = sheet
            .all_cells()
            .flat_map(|cell| rule.on_cell(&sheet, cell, &mut ctx))
            .collect();

        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_clean_text() {
        let sheet = make_sheet(vec![make_text_cell(0, 0, "hello")]);
        let rule = UnnecessarySpaceRule;
        let mut ctx = LinterContext::default();

        let violations: Vec<_> = sheet
            .all_cells()
            .flat_map(|cell| rule.on_cell(&sheet, cell, &mut ctx))
            .collect();

        assert!(violations.is_empty());
    }

    #[test]
    fn test_number_cell_ignored() {
        let mut cell_map = HashMap::new();
        cell_map.insert(
            (0, 0),
            Cell {
                row: 0,
                col: 0,
                value: CellValue::Number(42.0),
                formula: None,
                is_array: false,
                num_fmt: None,
            },
        );
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cell_map,
            ..Default::default()
        };
        let rule = UnnecessarySpaceRule;
        let mut ctx = LinterContext::default();

        let violations: Vec<_> = sheet
            .all_cells()
            .flat_map(|cell| rule.on_cell(&sheet, cell, &mut ctx))
            .collect();

        assert!(violations.is_empty());
    }

    #[test]
    fn test_empty_string() {
        let sheet = make_sheet(vec![make_text_cell(0, 0, "")]);
        let rule = UnnecessarySpaceRule;
        let mut ctx = LinterContext::default();

        let violations: Vec<_> = sheet
            .all_cells()
            .flat_map(|cell| rule.on_cell(&sheet, cell, &mut ctx))
            .collect();

        assert!(violations.is_empty());
    }
}

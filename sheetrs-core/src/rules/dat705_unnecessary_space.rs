//! DATA705: Unnecessary spaces
//!
//! Spaces before or after a text are invisible but change its value: `"North "` does not match
//! `"North"` in a lookup, a comparison or a filter, so formulas fail silently. The padding is
//! almost always an accident of copying or typing.
//!
//! # Reports
//!
//! One violation per cell (unit: the cell) whose text value begins or ends with an ASCII white-space
//! character (space, tab, line feed, form feed or carriage return). A text made only of white space
//! is reported. Text that a formula returns is checked like any other text value.
//!
//! # Does not report
//!
//! - Spaces inside the text, double spaces between words included.
//! - Empty text, and numbers, booleans, dates and errors: only text values are inspected.
//! - Unicode white space other than the ASCII characters above, such as the non-breaking space
//!   (U+00A0): the check looks at the first and last byte only.
//!
//! # Parameters
//!
//! None.
//!
//! # Severity
//!
//! Always `Info`: unlike most data rules the problem is cosmetic until a lookup depends on the text.

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

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        at: (u32, u32),
        cell: &Cell,
        _ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        if let crate::reader::CellValue::Text(text) = &cell.value {
            let bytes = text.as_bytes();
            let has_whitespace = bytes.first().is_some_and(|b| b.is_ascii_whitespace())
                || bytes.last().is_some_and(|b| b.is_ascii_whitespace());

            if has_whitespace {
                out.push(Violation::with_data(
                    RuleId::Data705,
                    ViolationScope::Cell(sheet.sheet_index, CellReference::new(at.0, at.1)),
                    UnnecessarySpaceData,
                    Severity::Info,
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
    use std::sync::Arc;

    fn make_text_cell(row: u32, col: u32, text: &str) -> ((u32, u32), Cell) {
        (
            (row, col),
            Cell {
                value: CellValue::Text(Arc::from(text)),
                formula: None,
                is_array: false,
                num_fmt: None,
            },
        )
    }

    fn make_sheet(cells: Vec<((u32, u32), Cell)>) -> Sheet {
        let mut cell_map = HashMap::new();
        for (at, c) in cells {
            cell_map.insert(at, c);
        }
        Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cell_map.into(),
            ..Default::default()
        }
    }

    #[test]
    fn test_leading_space() {
        let sheet = make_sheet(vec![make_text_cell(0, 0, " hello")]);
        let mut rule = UnnecessarySpaceRule;
        let mut ctx = LinterContext::default();

        let violations: Vec<_> = sheet
            .cells
            .iter()
            .flat_map(|(at, cell)| {
                collect_violations(|out| rule.on_cell(&sheet, at, cell, &mut ctx, out))
            })
            .collect();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Data705);
    }

    #[test]
    fn test_trailing_space() {
        let sheet = make_sheet(vec![make_text_cell(0, 0, "hello ")]);
        let mut rule = UnnecessarySpaceRule;
        let mut ctx = LinterContext::default();

        let violations: Vec<_> = sheet
            .cells
            .iter()
            .flat_map(|(at, cell)| {
                collect_violations(|out| rule.on_cell(&sheet, at, cell, &mut ctx, out))
            })
            .collect();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Data705);
    }

    #[test]
    fn test_both_spaces() {
        let sheet = make_sheet(vec![make_text_cell(0, 0, " hello ")]);
        let mut rule = UnnecessarySpaceRule;
        let mut ctx = LinterContext::default();

        let violations: Vec<_> = sheet
            .cells
            .iter()
            .flat_map(|(at, cell)| {
                collect_violations(|out| rule.on_cell(&sheet, at, cell, &mut ctx, out))
            })
            .collect();

        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_clean_text() {
        let sheet = make_sheet(vec![make_text_cell(0, 0, "hello")]);
        let mut rule = UnnecessarySpaceRule;
        let mut ctx = LinterContext::default();

        let violations: Vec<_> = sheet
            .cells
            .iter()
            .flat_map(|(at, cell)| {
                collect_violations(|out| rule.on_cell(&sheet, at, cell, &mut ctx, out))
            })
            .collect();

        assert!(violations.is_empty());
    }

    #[test]
    fn test_number_cell_ignored() {
        let mut cell_map = HashMap::new();
        cell_map.insert(
            (0, 0),
            Cell {
                value: CellValue::Number(42.0),
                formula: None,
                is_array: false,
                num_fmt: None,
            },
        );
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cell_map.into(),
            ..Default::default()
        };
        let mut rule = UnnecessarySpaceRule;
        let mut ctx = LinterContext::default();

        let violations: Vec<_> = sheet
            .cells
            .iter()
            .flat_map(|(at, cell)| {
                collect_violations(|out| rule.on_cell(&sheet, at, cell, &mut ctx, out))
            })
            .collect();

        assert!(violations.is_empty());
    }

    #[test]
    fn test_empty_string() {
        let sheet = make_sheet(vec![make_text_cell(0, 0, "")]);
        let mut rule = UnnecessarySpaceRule;
        let mut ctx = LinterContext::default();

        let violations: Vec<_> = sheet
            .cells
            .iter()
            .flat_map(|(at, cell)| {
                collect_violations(|out| rule.on_cell(&sheet, at, cell, &mut ctx, out))
            })
            .collect();

        assert!(violations.is_empty());
    }
}

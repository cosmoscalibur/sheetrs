//! VUL603: Empty string test detection
//!
//! Description: Recommends using ISBLANK() instead of checking for "" in formulas.

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use regex::Regex;
use std::collections::HashMap;
use std::sync::LazyLock;

/// Matches `=""` or `<>""` patterns.
static EMPTY_STRING_EQ_RE: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(r#"(=|<>)\s*"""#).expect("VUL603 empty string eq regex must compile")
});

/// Matches `LEN(...)=0` or `LEN(...)>0` patterns.
static LEN_ZERO_RE: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(r"LEN\s*\([^)]+\)\s*(=|<>|>|<)\s*0").expect("VUL603 LEN zero regex must compile")
});

/// Per-sheet cell collection: sheet_index → Vec<(row, col)>.
type SheetCellMap = HashMap<u16, Vec<(u32, u32)>>;

/// Rule that detects inefficient empty string tests in formulas.
pub struct EmptyStringTestRule {
    /// Per-sheet collected cells.
    sheet_cells: SheetCellMap,
}

impl EmptyStringTestRule {
    /// Create a new instance.
    pub fn new() -> Self {
        Self {
            sheet_cells: HashMap::new(),
        }
    }
}

impl Default for EmptyStringTestRule {
    fn default() -> Self {
        Self::new()
    }
}

/// Incident data for VUL603.
#[derive(Debug)]
pub struct EmptyStringTestData {
    /// Bounding box of violating cells (start_row, start_col, end_row, end_col).
    pub range: (u32, u32, u32, u32),
}

impl ViolationData for EmptyStringTestData {
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

        format!(
            "Empty string test (=\"\" or LEN()=0) found in range: {range_str}. Consider using ISBLANK() for better readability."
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for EmptyStringTestRule {
    fn id(&self) -> RuleId {
        RuleId::Vul603
    }

    fn name(&self) -> &'static str {
        "Empty String Test"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Vulnerability
    }

    fn on_cell(&mut self, sheet: &Sheet, cell: &Cell, ctx: &mut LinterContext) -> Vec<Violation> {
        if let Some(formula_upper) = ctx.formula_upper(sheet, cell)
            && (EMPTY_STRING_EQ_RE.is_match(formula_upper) || LEN_ZERO_RE.is_match(formula_upper))
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
                RuleId::Vul603,
                ViolationScope::Sheet(sheet.sheet_index),
                EmptyStringTestData { range: bbox },
                Severity::Info,
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

    #[test]
    fn test_empty_string_equals() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from(r#"=IF(A1="","Empty","Not Empty")"#)),
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

        let mut rule = EmptyStringTestRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Vul603);
    }

    #[test]
    fn test_len_equals_zero() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=IF(LEN(A1)=0,\"Empty\",\"Not Empty\")")),
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

        let mut rule = EmptyStringTestRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Vul603);
    }

    #[test]
    fn test_isblank_usage() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=IF(ISBLANK(A1),\"Empty\",\"Not Empty\")")),
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

        let mut rule = EmptyStringTestRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert!(violations.is_empty());
    }
}

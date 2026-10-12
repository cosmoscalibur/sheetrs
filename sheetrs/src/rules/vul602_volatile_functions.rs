//! VUL602: Avoid volatile functions
//!
//! Description: Volatile functions (like NOW, TODAY, RAND) recalculate every time
//! the sheet recalculates, potentially causing performance issues in large workbooks.

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::HashMap;
use std::sync::LazyLock;

/// Fixed list of volatile functions as internal constants.
const VOLATILE_FUNCTIONS: &[&str] = &[
    "NOW",
    "TODAY",
    "RAND",
    "RANDBETWEEN",
    "OFFSET",
    "INDIRECT",
    "INFO",
    "CELL",
];

/// Pre-computed search patterns: "FUNC(" for each volatile function.
static SEARCH_PATTERNS: LazyLock<Vec<String>> = LazyLock::new(|| {
    VOLATILE_FUNCTIONS
        .iter()
        .map(|name| format!("{name}("))
        .collect()
});

/// Per-sheet cell collection: sheet_index → Vec<(func_index, row, col)>.
type SheetCellMap = HashMap<u16, Vec<(u8, u32, u32)>>;

/// Rule that detects usage of volatile functions.
pub struct VolatileFunctionsRule {
    /// Per-sheet collected cells with function index.
    sheet_cells: SheetCellMap,
}

impl VolatileFunctionsRule {
    /// Create a new instance.
    pub fn new() -> Self {
        Self {
            sheet_cells: HashMap::new(),
        }
    }
}

impl Default for VolatileFunctionsRule {
    fn default() -> Self {
        Self::new()
    }
}

/// Incident data for VUL602.
#[derive(Debug)]
pub struct VolatileFunctionData {
    /// Index into `VOLATILE_FUNCTIONS` list.
    pub func_index: u8,
    /// Bounding box of violating cells (start_row, start_col, end_row, end_col).
    pub range: (u32, u32, u32, u32),
}

impl ViolationData for VolatileFunctionData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let func_name = VOLATILE_FUNCTIONS
            .get(self.func_index as usize)
            .unwrap_or(&"UNKNOWN");

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
            "Volatile function {func_name}() found in range: {range_str}. Consider alternatives for better performance."
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

impl WalkerRule for VolatileFunctionsRule {
    fn id(&self) -> RuleId {
        RuleId::Vul602
    }

    fn name(&self) -> &'static str {
        "Volatile Function"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Vulnerability
    }

    fn on_cell(&mut self, sheet: &Sheet, cell: &Cell, ctx: &mut LinterContext) -> Vec<Violation> {
        if let Some(formula_upper) = ctx.formula_upper(sheet, cell)
            && let Some(idx) = SEARCH_PATTERNS
                .iter()
                .position(|pattern| formula_upper.contains(pattern.as_str()))
        {
            self.sheet_cells
                .entry(sheet.sheet_index)
                .or_default()
                .push((idx as u8, cell.row, cell.col));
        }
        Vec::new()
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext) -> Vec<Violation> {
        let Some(cells) = self.sheet_cells.remove(&sheet.sheet_index) else {
            return Vec::new();
        };

        let mut violations = Vec::new();

        // Group by function index
        let mut by_func: HashMap<u8, Vec<(u32, u32)>> = HashMap::new();
        for (func_idx, row, col) in cells {
            by_func.entry(func_idx).or_default().push((row, col));
        }

        for (func_idx, func_cells) in by_func {
            let ranges = find_contiguous_ranges(&func_cells);
            for range in ranges {
                let bbox = bounding_box(&range);
                violations.push(Violation::with_data(
                    RuleId::Vul602,
                    ViolationScope::Sheet(sheet.sheet_index),
                    VolatileFunctionData {
                        func_index: func_idx,
                        range: bbox,
                    },
                    Severity::Info,
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
    fn test_volatile_function_now() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=NOW()")),
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

        let mut rule = VolatileFunctionsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Vul602);
    }

    #[test]
    fn test_multiple_volatile_functions() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=RAND()")),
                is_array: false,
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Empty,
            },
        );
        cells.insert(
            (1, 0),
            Cell {
                formula: Some(<Box<str>>::from("=TODAY()")),
                is_array: false,
                num_fmt: None,
                row: 1,
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

        let mut rule = VolatileFunctionsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert_eq!(violations.len(), 2);
    }

    #[test]
    fn test_case_insensitive() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: Some(<Box<str>>::from("=now()")),
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

        let mut rule = VolatileFunctionsRule::new();
        let mut ctx = LinterContext::default();
        for cell in sheet.all_cells() {
            rule.on_cell(&sheet, cell, &mut ctx);
        }
        let violations = rule.on_sheet_end(&sheet, &mut ctx);

        assert_eq!(violations.len(), 1);
    }
}

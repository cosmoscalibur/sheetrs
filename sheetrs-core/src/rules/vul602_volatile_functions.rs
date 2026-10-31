//! VUL602: Volatile functions
//!
//! A volatile function is recalculated every time anything in the workbook is recalculated, together
//! with all the cells that depend on it. Many of them slow a large workbook down, and `NOW`, `TODAY`
//! and `RAND` make results change without anybody editing the file.
//!
//! # Reports
//! One violation per group of adjacent cells (scope: sheet) that use the same volatile function; the
//! message names the function and gives the group's bounding box. The functions are a fixed list:
//! `NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `OFFSET`, `INDIRECT`, `INFO` and `CELL`. They are found as
//! a call `NAME(` in the upper-cased formula (lower case is accepted) outside string literals and
//! quoted sheet names, and not as the end of a longer or prefixed name (`MYNOW(` is not `NOW(`). A
//! cell that calls several of them is placed only under the first one of the list above, which is
//! also the order in which a cell is attributed.
//!
//! # Does not report
//! Any function not in the list, including functions that are volatile only in some contexts.
//! Whether the volatile result is used in a critical calculation is not examined.
//!
//! # Parameters
//! None: the list is fixed in the code and cannot be changed in the configuration.
//!
//! # Severity
//! Always `Info`.

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::helpers::{has_call, mask_quoted};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::{BTreeMap, HashMap};
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

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        at: (u32, u32),
        cell: &Cell,
        ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        if let Some(formula_upper) = ctx.formula_upper(sheet, at, cell) {
            // A cheap search of the text first; a candidate is confirmed as a call of the
            // function (not the end of a longer name, not inside a string literal).
            let mut masked = None;
            let found = SEARCH_PATTERNS.iter().position(|pattern| {
                formula_upper.contains(pattern.as_str()) && {
                    let m = masked.get_or_insert_with(|| mask_quoted(formula_upper));
                    has_call(m, &pattern[..pattern.len() - 1])
                }
            });
            let Some(idx) = found else {
                return;
            };
            self.sheet_cells
                .entry(sheet.sheet_index)
                .or_default()
                .push((idx as u8, at.0, at.1));
        }
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext, out: &mut Vec<Violation>) {
        let Some(mut cells) = self.sheet_cells.remove(&sheet.sheet_index) else {
            return;
        };
        cells.sort_unstable();

        // Group by function index
        let mut by_func: BTreeMap<u8, Vec<(u32, u32)>> = BTreeMap::new();
        for (func_idx, row, col) in cells {
            by_func.entry(func_idx).or_default().push((row, col));
        }

        for (func_idx, func_cells) in by_func {
            let ranges = find_contiguous_ranges(&func_cells);
            for range in ranges {
                let bbox = bounding_box(&range);
                out.push(Violation::with_data(
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
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::collect_violations;
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
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            ..Default::default()
        };

        let mut rule = VolatileFunctionsRule::new();
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

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
                value: CellValue::Empty,
            },
        );
        cells.insert(
            (1, 0),
            Cell {
                formula: Some(<Box<str>>::from("=TODAY()")),
                is_array: false,
                num_fmt: None,
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            ..Default::default()
        };

        let mut rule = VolatileFunctionsRule::new();
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

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
                value: CellValue::Empty,
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            ..Default::default()
        };

        let mut rule = VolatileFunctionsRule::new();
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_violations_ordered_by_function_and_cell() {
        let formulas = ["=NOW()", "=RAND()", "=TODAY()"];
        let mut cells = HashMap::new();
        for i in 0..18u32 {
            let (row, col) = (i * 2, 0);
            cells.insert(
                (row, col),
                Cell {
                    formula: Some(<Box<str>>::from(formulas[i as usize % 3])),
                    is_array: false,
                    num_fmt: None,
                    value: CellValue::Empty,
                },
            );
        }
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            ..Default::default()
        };

        let mut rule = VolatileFunctionsRule::new();
        let mut ctx = LinterContext::default();
        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }
        let keys: Vec<_> = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out))
            .iter()
            .map(|v| {
                let data = v.data::<VolatileFunctionData>().unwrap();
                (data.func_index, data.range)
            })
            .collect();
        let mut sorted = keys.clone();
        sorted.sort();
        assert_eq!(keys.len(), 18);
        assert_eq!(keys, sorted);
    }

    #[test]
    fn a_volatile_name_is_a_call_not_text_or_the_end_of_another_name() {
        let masked = |f: &str| mask_quoted(f).into_owned();
        assert!(has_call(&masked("=A1+NOW()"), "NOW"));
        assert!(!has_call(&masked("=\"NOW(\"&A1"), "NOW"));
        assert!(!has_call(&masked("=MYNOW()"), "NOW"));
        assert!(!has_call(&masked("='NOW(' !A1"), "NOW"));
        assert!(has_call(&masked("=IF(A1,\"x\",RAND())"), "RAND"));
    }
}

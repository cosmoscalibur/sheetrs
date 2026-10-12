//! REF310: Longer Ref Expected detection
//!
//! Description: Identifies cases where a formula's range reference stops
//! abruptly before adjacent data, suggesting an incomplete or too-short range.
//!
//! **Dependency:** requires CALC202 (Circular References) to be active,
//! as this rule re-uses the same `CELL_REF_PATTERN` regex to extract range
//! references from formulas.

use super::helpers::{CELL_REF_PATTERN, parse_cell_coords};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::reader::workbook::CellValue;
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::{HashMap, HashSet};

/// Rule that identifies suspiciously short references.
pub struct LongerRefExpectedRule {
    /// Minimum number of adjacent data cells beyond the range end to trigger.
    min_adjacent_data: usize,
}

impl LongerRefExpectedRule {
    pub fn new(config: &LinterConfig) -> Self {
        Self {
            min_adjacent_data: config
                .get_param_int("ref310_min_adjacent_data", None)
                .unwrap_or(1) as usize,
        }
    }
}

/// Incident data for REF310.
#[derive(Debug)]
pub struct LongerRefData {
    /// The short range reference (display string).
    pub range_str: String,
    /// Location of the adjacent data cell.
    pub adjacent_cell: (u16, u32, u32),
}

impl ViolationData for LongerRefData {
    fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        let (sheet_idx, row, col) = self.adjacent_cell;
        let sheet_name = ctx
            .workbook
            .sheets
            .get(sheet_idx as usize)
            .map(|s| s.name.as_str())
            .unwrap_or("?");
        format!(
            "Range {} may be too short — adjacent cell {}!{} contains data.",
            self.range_str,
            sheet_name,
            CellReference::new(row, col)
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Determine the value-type category of a cell for type-matching suppression.
/// Returns `None` for empty cells, or a category discriminant for non-empty.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum ValueCategory {
    Numeric,
    Text,
    Boolean,
    Error,
}

fn categorize_value(value: &CellValue) -> Option<ValueCategory> {
    match value {
        CellValue::Empty => None,
        CellValue::Number(_) => Some(ValueCategory::Numeric),
        CellValue::Text(_) => Some(ValueCategory::Text),
        CellValue::Boolean(_) => Some(ValueCategory::Boolean),
        CellValue::Error(_) => Some(ValueCategory::Error),
    }
}

/// Determine the dominant value category of cells within a range on a sheet.
fn dominant_category(
    cells: &HashMap<(u32, u32), crate::reader::Cell>,
    start_row: u32,
    start_col: u32,
    end_row: u32,
    end_col: u32,
) -> Option<ValueCategory> {
    let mut counts: HashMap<ValueCategory, usize> = HashMap::new();
    for row in start_row..=end_row {
        for col in start_col..=end_col {
            if let Some(cell) = cells.get(&(row, col))
                && let Some(cat) = categorize_value(&cell.value)
            {
                *counts.entry(cat).or_default() += 1;
            }
        }
    }
    counts
        .into_iter()
        .max_by_key(|(_, count)| *count)
        .map(|(cat, _)| cat)
}

/// A range reference extracted from a formula.
struct RangeRef {
    /// Target sheet index.
    sheet_index: u16,
    /// Range coordinates (start_row, start_col, end_row, end_col).
    start_row: u32,
    start_col: u32,
    end_row: u32,
    end_col: u32,
    /// Original range string for violation messages.
    display: String,
}

/// Extract range references from a formula (only ranges, not single-cell refs).
fn extract_range_refs(
    formula: &str,
    current_sheet_index: u16,
    name_to_index: &HashMap<String, u16>,
) -> Vec<RangeRef> {
    let mut ranges = Vec::new();

    for cap in CELL_REF_PATTERN.captures_iter(formula) {
        // Only process range references (those with end col/row groups)
        let (Some(end_col_match), Some(end_row_match)) = (cap.get(6), cap.get(7)) else {
            continue; // Single cell ref — skip
        };

        // Resolve sheet index
        let sheet_index = if cap.get(1).is_some() {
            if let Some(quoted) = cap.get(2) {
                match name_to_index.get(quoted.as_str()).copied() {
                    Some(idx) => idx,
                    None => continue,
                }
            } else if let Some(unquoted) = cap.get(3) {
                match name_to_index.get(unquoted.as_str()).copied() {
                    Some(idx) => idx,
                    None => continue,
                }
            } else {
                current_sheet_index
            }
        } else {
            current_sheet_index
        };

        // Parse start coordinates
        let (Some(c_match), Some(r_match)) = (cap.get(4), cap.get(5)) else {
            continue;
        };
        let Some((start_row, start_col)) = parse_cell_coords(r_match.as_str(), c_match.as_str())
        else {
            continue;
        };

        // Parse end coordinates
        let Some((end_row, end_col)) =
            parse_cell_coords(end_row_match.as_str(), end_col_match.as_str())
        else {
            continue;
        };

        let display = cap.get(0).map_or("", |m| m.as_str()).to_string();

        ranges.push(RangeRef {
            sheet_index,
            start_row,
            start_col,
            end_row,
            end_col,
            display,
        });
    }

    ranges
}

impl WalkerRule for LongerRefExpectedRule {
    fn id(&self) -> RuleId {
        RuleId::Ref310
    }

    fn name(&self) -> &'static str {
        "Longer Ref Expected"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::Reference
    }

    fn on_workbook_end(&mut self, workbook: &Workbook, ctx: &mut LinterContext) -> Vec<Violation> {
        if ctx.cell_dependencies.is_empty() {
            return Vec::new();
        }

        // Build sheet lookup
        let sheets_by_index: HashMap<u16, &crate::reader::Sheet> =
            workbook.sheets.iter().map(|s| (s.sheet_index, s)).collect();

        let mut violations = Vec::new();
        // Track which cells we've already flagged to avoid duplicates
        let mut flagged: HashSet<(u16, u32, u32)> = HashSet::new();

        // Iterate through all formula cells (keys of cell_dependencies)
        for formula_loc in ctx.cell_dependencies.keys() {
            let (formula_sheet_idx, formula_row, formula_col) = *formula_loc;

            // Get the formula text
            let Some(formula) = sheets_by_index
                .get(&formula_sheet_idx)
                .and_then(|s| s.get_cell(formula_row, formula_col))
                .and_then(|c| c.as_formula())
            else {
                continue;
            };
            let formula = formula.to_string();

            // Extract range references from this formula
            let range_refs = extract_range_refs(&formula, formula_sheet_idx, &ctx.name_to_index);

            for range_ref in &range_refs {
                let Some(target_sheet) = sheets_by_index.get(&range_ref.sheet_index) else {
                    continue;
                };

                // Determine the dominant axis: vertical if range spans more rows than cols
                let row_span = range_ref.end_row.saturating_sub(range_ref.start_row) + 1;
                let col_span = range_ref.end_col.saturating_sub(range_ref.start_col) + 1;

                // Check vertical extension (row after end_row)
                if row_span >= col_span {
                    let adjacent_row = range_ref.end_row + 1;
                    let mut adjacent_count = 0usize;

                    for check_row in adjacent_row.. {
                        let mut has_data = false;
                        for col in range_ref.start_col..=range_ref.end_col {
                            if let Some(cell) = target_sheet.get_cell(check_row, col)
                                && (!cell.value.is_empty() || cell.is_formula())
                            {
                                has_data = true;
                                break;
                            }
                        }
                        if has_data {
                            adjacent_count += 1;
                        } else {
                            break;
                        }

                        // Stop checking after we've confirmed enough
                        if adjacent_count >= self.min_adjacent_data {
                            break;
                        }
                    }

                    if adjacent_count >= self.min_adjacent_data {
                        let adj_cell = (range_ref.sheet_index, adjacent_row, range_ref.start_col);
                        if flagged.insert(*formula_loc) {
                            // Type-matching suppression: skip if adjacent data
                            // has a different dominant type than the range
                            let range_cat = dominant_category(
                                &target_sheet.cells,
                                range_ref.start_row,
                                range_ref.start_col,
                                range_ref.end_row,
                                range_ref.end_col,
                            );
                            let adj_cat = target_sheet
                                .get_cell(adjacent_row, range_ref.start_col)
                                .and_then(|c| categorize_value(&c.value));

                            if let (Some(rc), Some(ac)) = (range_cat, adj_cat)
                                && rc != ac
                            {
                                continue; // Different type → likely a header
                            }

                            violations.push(Violation::with_data(
                                RuleId::Ref310,
                                ViolationScope::Cell(
                                    formula_sheet_idx,
                                    CellReference::new(formula_row, formula_col),
                                ),
                                LongerRefData {
                                    range_str: range_ref.display.clone(),
                                    adjacent_cell: adj_cell,
                                },
                                Severity::Info,
                            ));
                        }
                    }
                }

                // Check horizontal extension (col after end_col)
                if col_span > row_span {
                    let adjacent_col = range_ref.end_col + 1;
                    let mut adjacent_count = 0usize;

                    for check_col in adjacent_col.. {
                        let mut has_data = false;
                        for row in range_ref.start_row..=range_ref.end_row {
                            if let Some(cell) = target_sheet.get_cell(row, check_col)
                                && (!cell.value.is_empty() || cell.is_formula())
                            {
                                has_data = true;
                                break;
                            }
                        }
                        if has_data {
                            adjacent_count += 1;
                        } else {
                            break;
                        }

                        if adjacent_count >= self.min_adjacent_data {
                            break;
                        }
                    }

                    if adjacent_count >= self.min_adjacent_data {
                        let adj_cell = (range_ref.sheet_index, range_ref.start_row, adjacent_col);
                        if flagged.insert(*formula_loc) {
                            let range_cat = dominant_category(
                                &target_sheet.cells,
                                range_ref.start_row,
                                range_ref.start_col,
                                range_ref.end_row,
                                range_ref.end_col,
                            );
                            let adj_cat = target_sheet
                                .get_cell(range_ref.start_row, adjacent_col)
                                .and_then(|c| categorize_value(&c.value));

                            if let (Some(rc), Some(ac)) = (range_cat, adj_cat)
                                && rc != ac
                            {
                                continue;
                            }

                            violations.push(Violation::with_data(
                                RuleId::Ref310,
                                ViolationScope::Cell(
                                    formula_sheet_idx,
                                    CellReference::new(formula_row, formula_col),
                                ),
                                LongerRefData {
                                    range_str: range_ref.display.clone(),
                                    adjacent_cell: adj_cell,
                                },
                                Severity::Info,
                            ));
                        }
                    }
                }
            }
        }

        violations
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet, Workbook};
    use crate::rules::LinterContext;
    use std::collections::HashMap;
    use std::sync::Arc;

    fn make_number_cell(row: u32, col: u32, n: f64) -> Cell {
        Cell {
            row,
            col,
            value: CellValue::Number(n),
            formula: None,
            is_array: false,
            num_fmt: None,
        }
    }

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

    fn make_formula_cell(row: u32, col: u32, formula: &str) -> Cell {
        Cell {
            row,
            col,
            value: CellValue::Empty,
            formula: Some(Box::from(formula)),
            is_array: false,
            num_fmt: None,
        }
    }

    fn make_sheet_with_index(index: u16, name: &str, cells: Vec<Cell>) -> Sheet {
        let mut cell_map = HashMap::new();
        for cell in cells {
            cell_map.insert((cell.row, cell.col), cell);
        }
        Sheet {
            name: name.to_string(),
            sheet_index: index,
            cells: cell_map,
            ..Default::default()
        }
    }

    fn make_workbook(sheets: Vec<Sheet>) -> Workbook {
        Workbook {
            sheets,
            ..Default::default()
        }
    }

    #[test]
    fn test_sum_range_with_adjacent_data() {
        // A1:A5 have numbers, A6 also has a number → flag
        let mut cells = Vec::new();
        for row in 0..6 {
            cells.push(make_number_cell(row, 0, (row + 1) as f64));
        }
        // B1 = SUM(A1:A5)
        cells.push(make_formula_cell(0, 1, "SUM(A1:A5)"));

        let sheet = make_sheet_with_index(0, "Sheet1", cells);
        let workbook = make_workbook(vec![sheet]);

        let config = LinterConfig::default();
        let mut rule = LongerRefExpectedRule::new(&config);
        let mut ctx = LinterContext::default();
        ctx.name_to_index.insert("Sheet1".to_string(), 0);
        // B1 depends on A1..A5
        ctx.cell_dependencies
            .insert((0, 0, 1), vec![(0, 0, 0), (0, 4, 0)]);

        let violations = rule.on_workbook_end(&workbook, &mut ctx);
        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ref310);
    }

    #[test]
    fn test_sum_range_no_adjacent_data() {
        // A1:A5 have numbers, A6 is empty → no flag
        let mut cells = Vec::new();
        for row in 0..5 {
            cells.push(make_number_cell(row, 0, (row + 1) as f64));
        }
        cells.push(make_formula_cell(0, 1, "SUM(A1:A5)"));

        let sheet = make_sheet_with_index(0, "Sheet1", cells);
        let workbook = make_workbook(vec![sheet]);

        let config = LinterConfig::default();
        let mut rule = LongerRefExpectedRule::new(&config);
        let mut ctx = LinterContext::default();
        ctx.name_to_index.insert("Sheet1".to_string(), 0);
        ctx.cell_dependencies
            .insert((0, 0, 1), vec![(0, 0, 0), (0, 4, 0)]);

        let violations = rule.on_workbook_end(&workbook, &mut ctx);
        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_type_mismatch_suppression() {
        // A1:A5 have numbers, A6 is text (header) → suppress
        let mut cells = Vec::new();
        for row in 0..5 {
            cells.push(make_number_cell(row, 0, (row + 1) as f64));
        }
        cells.push(make_text_cell(5, 0, "Total"));
        cells.push(make_formula_cell(0, 1, "SUM(A1:A5)"));

        let sheet = make_sheet_with_index(0, "Sheet1", cells);
        let workbook = make_workbook(vec![sheet]);

        let config = LinterConfig::default();
        let mut rule = LongerRefExpectedRule::new(&config);
        let mut ctx = LinterContext::default();
        ctx.name_to_index.insert("Sheet1".to_string(), 0);
        ctx.cell_dependencies
            .insert((0, 0, 1), vec![(0, 0, 0), (0, 4, 0)]);

        let violations = rule.on_workbook_end(&workbook, &mut ctx);
        assert_eq!(
            violations.len(),
            0,
            "Should suppress when adjacent cell is text but range is numeric"
        );
    }

    #[test]
    fn test_horizontal_range_with_adjacent_data() {
        // B1:E1 have numbers, F1 also has a number → flag
        let mut cells = Vec::new();
        for col in 1..6 {
            cells.push(make_number_cell(0, col, col as f64));
        }
        // A1 = SUM(B1:E1)
        cells.push(make_formula_cell(0, 0, "SUM(B1:E1)"));

        let sheet = make_sheet_with_index(0, "Sheet1", cells);
        let workbook = make_workbook(vec![sheet]);

        let config = LinterConfig::default();
        let mut rule = LongerRefExpectedRule::new(&config);
        let mut ctx = LinterContext::default();
        ctx.name_to_index.insert("Sheet1".to_string(), 0);
        ctx.cell_dependencies
            .insert((0, 0, 0), vec![(0, 0, 1), (0, 0, 4)]);

        let violations = rule.on_workbook_end(&workbook, &mut ctx);
        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_no_formula() {
        let sheet = make_sheet_with_index(0, "Sheet1", vec![make_number_cell(0, 0, 42.0)]);
        let workbook = make_workbook(vec![sheet]);

        let config = LinterConfig::default();
        let mut rule = LongerRefExpectedRule::new(&config);
        let ctx = &mut LinterContext::default();

        let violations = rule.on_workbook_end(&workbook, ctx);
        assert_eq!(violations.len(), 0);
    }

    #[test]
    fn test_config_min_adjacent_data() {
        // A1:A5 numbers, A6 has number but min_adjacent_data=2, A7 is empty → no flag
        let mut cells = Vec::new();
        for row in 0..6 {
            cells.push(make_number_cell(row, 0, (row + 1) as f64));
        }
        cells.push(make_formula_cell(0, 1, "SUM(A1:A5)"));

        let sheet = make_sheet_with_index(0, "Sheet1", cells);
        let workbook = make_workbook(vec![sheet]);

        let mut config = LinterConfig::default();
        config.global.params.insert(
            "ref310_min_adjacent_data".to_string(),
            toml::Value::Integer(2),
        );
        let mut rule = LongerRefExpectedRule::new(&config);
        let mut ctx = LinterContext::default();
        ctx.name_to_index.insert("Sheet1".to_string(), 0);
        ctx.cell_dependencies
            .insert((0, 0, 1), vec![(0, 0, 0), (0, 4, 0)]);

        let violations = rule.on_workbook_end(&workbook, &mut ctx);
        assert_eq!(
            violations.len(),
            0,
            "min_adjacent_data=2 but only 1 adjacent → no flag"
        );
    }
}

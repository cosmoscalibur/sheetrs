//! REF310: Longer reference expected
//!
//! A range in a formula that stops right before more data of the same kind, such as
//! `=SUM(A1:A5)` with values still in A6. When rows or columns are added after the formula was
//! written, the range silently leaves them out and the result is wrong without any error.
//!
//! # Reports
//!
//! One violation per formula cell and qualifying range (the unit is the cell). Every distinct
//! range of the formula is examined, on its own sheet when it is qualified with a sheet name. The
//! direction is vertical when the range has at least as many rows as columns, otherwise
//! horizontal; only that direction is checked. A range qualifies when:
//!
//! - the line right after it (the row below, or the column to the right, over the range's width
//!   or height) continues the block in any of its cells, and the following lines do too until
//!   `ref310_min_adjacent_data` consecutive lines are confirmed. A cell continues the block when
//!   the range's last cell in that position is a constant and it is a non-empty constant, or when
//!   the range's last cell is a formula and it is a formula of the same pattern (the same formula
//!   read relative to each cell, as INT401 compares formulas); and
//! - type matching holds: each cell of the range's last row (or column) has the same value
//!   category (number, text, boolean, error, empty) as its neighbour on that first adjacent line.
//!   A formula without a saved value matches any category. A mismatch, such as a text header
//!   below numbers, suggests a different block.
//!
//! The finding tells the whole block the range lacks, not only the next cell: the block goes on
//! while each following line continues it (as above, in any cell of the line), and the message
//! gives the part it adds with the sheet named and the replacement of the whole range, written like
//! the original (same `$` flags and sheet qualifier). For `=SUM(INGRESOS!$H$61:$Q$61)` with the
//! pattern going on to column AA: "Range INGRESOS!$H$61:$Q$61 may be too short: the block continues
//! in INGRESOS!$R$61:$AA$61 (10 columns). Suggested range: INGRESOS!$H$61:$AA$61." The same range
//! used by several formulas is examined once.
//!
//! The list of formula cells and the sheet names come from CALC202
//! ([`super::calc202_circular_references`], which must be active; without it nothing is
//! reported). The ranges themselves are found by a textual match of the formula with the shared
//! cell-reference pattern, not by the parsed formula, so a reference-like text inside a string
//! literal is also examined.
//!
//! # Does not report
//!
//! - Single-cell references, and whole-column or whole-row references (REF307).
//! - Ranges to sheets whose name is unknown.
//! - A different kind of cell next to the block: a formula under summed constants is a result (a
//!   total or subtotal), and constants under summed formulas are inputs. The same applies to a
//!   formula with another pattern (`=SUM(B2:B10)` over `=B2*2` rows, with `=SUM(B2:B10)` below).
//! - Adjacent data of another type than the range end, as above.
//! - Data separated from the range by a blank line: only the first line after the range is
//!   inspected.
//!
//! # Parameters
//!
//! - `ref310_min_adjacent_data` (integer of at least 1, default `1`, as in PerfectXL): the number of
//!   consecutive adjacent lines that continue the block needed to report. A configuration with a
//!   lower or non-integer value is refused when it is loaded.
//!
//! # Severity
//!
//! Always info.

use super::helpers::{CELL_REF_PATTERN, parse_cell_coords};
use super::int401_interrupted_by_data::InterruptionRule;
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::reader::workbook::CellValue;
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use std::collections::HashMap;
use std::hash::{BuildHasherDefault, Hasher};

/// Consecutive adjacent lines of data that make a range suspect, when the configuration does not
/// say.
const DEFAULT_MIN_ADJACENT_DATA: usize = 1;

/// Hasher of a key made of a few integers (a cell address): one multiplication per word. The
/// standard hasher is built to resist crafted keys, which the addresses of the workbook being
/// read cannot be.
#[derive(Default)]
struct AddressHasher(u64);

impl Hasher for AddressHasher {
    fn write(&mut self, bytes: &[u8]) {
        for &byte in bytes {
            self.write_u64(u64::from(byte));
        }
    }

    fn write_u16(&mut self, word: u16) {
        self.write_u64(u64::from(word));
    }

    fn write_u32(&mut self, word: u32) {
        self.write_u64(u64::from(word));
    }

    fn write_u64(&mut self, word: u64) {
        self.0 = (self.0.rotate_left(5) ^ word).wrapping_mul(0x517c_c1b7_2722_0a95);
    }

    /// The multiplication leaves its best mixed bits at the top and the table indexes by the
    /// bottom ones.
    fn finish(&self) -> u64 {
        self.0.rotate_left(26)
    }
}

/// Ids of the formula patterns of the cells compared while extending blocks. A cell is compared
/// by many ranges (each one ending at a different row of the same block), so its pattern is
/// computed once, and each distinct pattern is stored once.
#[derive(Default)]
struct PatternIds {
    by_cell: HashMap<(u16, u32, u32), u32, BuildHasherDefault<AddressHasher>>,
    interned: HashMap<Box<str>, u32>,
}

impl PatternIds {
    /// Returns the id of the pattern of `formula`, the formula of the cell at `at`.
    fn of(&mut self, sheet: &crate::reader::Sheet, at: (u32, u32), formula: &str) -> u32 {
        let key = (sheet.sheet_index, at.0, at.1);
        if let Some(&id) = self.by_cell.get(&key) {
            return id;
        }
        let pattern = InterruptionRule::normalize_formula(formula, at.0, at.1);
        let next = self.interned.len() as u32;
        let id = *self.interned.entry(pattern).or_insert(next);
        self.by_cell.insert(key, id);
        id
    }
}

/// Whether the cell at `candidate` continues the block that ends at `source`: a constant after a
/// constant, or a formula with the same pattern (the same formula once the references are read
/// relative to each cell) after a formula. A formula next to constants is a result (a total, a
/// subtotal) and constants next to a block of formulas are inputs, so neither extends the block.
fn continues_block(
    sheet: &crate::reader::Sheet,
    patterns: &mut PatternIds,
    source: (u32, u32),
    candidate: (u32, u32),
) -> bool {
    let (Some(src), Some(cand)) = (
        sheet.get_cell(source.0, source.1),
        sheet.get_cell(candidate.0, candidate.1),
    ) else {
        return false;
    };
    match (src.as_formula(), cand.as_formula()) {
        (None, None) => !src.value.is_empty() && !cand.value.is_empty(),
        (Some(f), Some(g)) => patterns.of(sheet, source, f) == patterns.of(sheet, candidate, g),
        _ => false,
    }
}

/// Rule that identifies suspiciously short references.
pub struct LongerRefExpectedRule {
    /// Minimum number of adjacent data cells beyond the range end to trigger.
    min_adjacent_data: usize,
}

impl LongerRefExpectedRule {
    /// Create the rule, reading `ref310_min_adjacent_data` from the configuration.
    pub fn new(config: &LinterConfig) -> Self {
        Self {
            // A value below 1 is refused when the configuration is validated; a configuration
            // built by hand gets the default instead.
            min_adjacent_data: config
                .get_param_int("ref310_min_adjacent_data", None)
                .filter(|n| *n >= 1)
                .map_or(DEFAULT_MIN_ADJACENT_DATA, |n| n as usize),
        }
    }
}

/// Incident data for REF310.
#[derive(Debug)]
pub struct LongerRefData {
    /// The short range reference, as the formula wrote it.
    pub range_str: String,
    /// The range extended over the whole block, written like the original.
    pub replacement: String,
    /// The part the range lacks, with the sheet named.
    pub added: String,
    /// Rows (vertical block) or columns (horizontal block) the range lacks.
    pub lines: usize,
    /// The block continues downwards (rows) when true, to the right (columns) otherwise.
    pub vertical: bool,
}

impl ViolationData for LongerRefData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let unit = match (self.vertical, self.lines) {
            (true, 1) => "row",
            (true, _) => "rows",
            (false, 1) => "column",
            (false, _) => "columns",
        };
        format!(
            "Range {} may be too short: the block continues in {} ({} {unit}). Suggested range: {}.",
            self.range_str, self.added, self.lines, self.replacement
        )
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Value-type category of a non-empty cell, for type-matching suppression.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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

/// Checks that each `(last, adjacent)` cell pair has the same value category,
/// treating an empty cell as its own category. A formula without a cached
/// value has an unknown type and matches any category.
fn boundary_types_match(
    sheet: &crate::reader::Sheet,
    mut pairs: impl Iterator<Item = ((u32, u32), (u32, u32))>,
) -> bool {
    let category = |(row, col)| {
        sheet
            .get_cell(row, col)
            .and_then(|c| categorize_value(&c.value))
    };
    let is_uncached_formula = |(row, col)| {
        sheet
            .get_cell(row, col)
            .is_some_and(|c| c.is_formula() && c.value.is_empty())
    };
    pairs.all(|(last, adjacent)| {
        is_uncached_formula(last)
            || is_uncached_formula(adjacent)
            || category(last) == category(adjacent)
    })
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

/// Extracts the distinct range references of a formula, by coordinates;
/// single-cell refs are skipped.
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

        let coords = (sheet_index, start_row, start_col, end_row, end_col);
        if ranges.iter().any(|r: &RangeRef| {
            (
                r.sheet_index,
                r.start_row,
                r.start_col,
                r.end_row,
                r.end_col,
            ) == coords
        }) {
            continue;
        }

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

    fn on_workbook_end(
        &mut self,
        workbook: &Workbook,
        ctx: &mut LinterContext,
        out: &mut Vec<Violation>,
    ) {
        if ctx.cell_dependencies.is_empty() {
            return;
        }

        // Build sheet lookup
        let sheets_by_index: HashMap<u16, &crate::reader::Sheet> =
            workbook.sheets.iter().map(|s| (s.sheet_index, s)).collect();

        // Extent of the block of each range, which many formulas share.
        let mut extents: HashMap<(u16, u32, u32, u32, u32), usize> = HashMap::new();
        let mut patterns = PatternIds::default();

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

                // The direction is vertical when the range has at least as many rows as columns.
                let vertical = (range_ref.end_row - range_ref.start_row)
                    >= (range_ref.end_col - range_ref.start_col);
                let key = (
                    range_ref.sheet_index,
                    range_ref.start_row,
                    range_ref.start_col,
                    range_ref.end_row,
                    range_ref.end_col,
                );
                let lines = *extents.entry(key).or_insert_with(|| {
                    block_extent(
                        target_sheet,
                        &mut patterns,
                        range_ref,
                        vertical,
                        self.min_adjacent_data,
                    )
                });
                if lines == 0 {
                    continue;
                }
                let sheet_name = target_sheet.name.as_str();
                let (replacement, added) =
                    suggestion(&range_ref.display, sheet_name, vertical, lines);
                out.push(Violation::with_data(
                    RuleId::Ref310,
                    ViolationScope::Cell(
                        formula_sheet_idx,
                        CellReference::new(formula_row, formula_col),
                    ),
                    LongerRefData {
                        range_str: range_ref.display.clone(),
                        replacement,
                        added,
                        lines,
                        vertical,
                    },
                    Severity::Info,
                ));
            }
        }
    }
}

/// How many lines (rows when `vertical`, columns otherwise) after the range continue its block,
/// or 0 when the range does not qualify: fewer than `min_lines` continue it, or the cells at its
/// end do not have the type of their neighbours on the first line after it.
fn block_extent(
    sheet: &crate::reader::Sheet,
    patterns: &mut PatternIds,
    range: &RangeRef,
    vertical: bool,
    min_lines: usize,
) -> usize {
    // Position `i` along the line, and the cell of the range's last line / of line `n` after it.
    let (first, last) = if vertical {
        (range.start_col, range.end_col)
    } else {
        (range.start_row, range.end_row)
    };
    let end = if vertical {
        range.end_row
    } else {
        range.end_col
    };
    let at = |i: u32, line: u32| if vertical { (line, i) } else { (i, line) };
    let mut lines = 0usize;
    loop {
        let line = end + 1 + lines as u32;
        if !(first..=last).any(|i| continues_block(sheet, patterns, at(i, end), at(i, line))) {
            break;
        }
        lines += 1;
    }
    if lines >= min_lines
        && boundary_types_match(sheet, (first..=last).map(|i| (at(i, end), at(i, end + 1))))
    {
        lines
    } else {
        0
    }
}

/// A cell reference as the formula wrote it: `$` flags, column letters, `$` flag, row digits.
fn split_cell_text(text: &str) -> Option<(bool, &str, bool, &str)> {
    let (col_abs, rest) = text.strip_prefix('$').map_or((false, text), |r| (true, r));
    let letters = rest.find(|c: char| !c.is_ascii_alphabetic())?;
    let (col, rest) = rest.split_at(letters);
    let (row_abs, row) = rest.strip_prefix('$').map_or((false, rest), |r| (true, r));
    (!col.is_empty() && !row.is_empty()).then_some((col_abs, col, row_abs, row))
}

/// The replacement of a range that should be longer, written as the range was (`$` flags and sheet
/// qualifier kept), and the part it adds with the sheet always named. The range is extended by
/// `lines` after its end in the direction of the block. A range this cannot read as `start:end`
/// (it came from the cell reference pattern, so it does) is returned unchanged.
fn suggestion(display: &str, sheet_name: &str, vertical: bool, lines: usize) -> (String, String) {
    let unchanged = || (display.to_string(), display.to_string());
    let Some((head, end_text)) = display.rsplit_once(':') else {
        return unchanged();
    };
    let (qualifier, start_text) = head
        .rsplit_once('!')
        .map_or(("", head), |(q, s)| (&head[..q.len() + 1], s));
    let (Some((sca, sc, sra, sr)), Some((eca, ec, era, er))) =
        (split_cell_text(start_text), split_cell_text(end_text))
    else {
        return unchanged();
    };
    let (Some(start_col), Ok(start_row), Some(end_col), Ok(end_row)) = (
        crate::reader::parser_utils::column_letters_to_index(sc),
        sr.parse::<u32>(),
        crate::reader::parser_utils::column_letters_to_index(ec),
        er.parse::<u32>(),
    ) else {
        return unchanged();
    };
    let n = lines as u32;
    let flag = |abs: bool| if abs { "$" } else { "" };
    // The first cell the range lacks (the block keeps the range's width or height) and the new
    // end, in the direction of the block.
    let (first_col, first_row, new_col, new_row) = if vertical {
        (start_col, end_row + 1, end_col, end_row + n)
    } else {
        (end_col + 1, start_row, end_col + n, end_row)
    };
    let letters = |col: u32| crate::reader::parser_utils::column_letters(col).to_string();
    let new_end = format!("{}{}{}{}", flag(eca), letters(new_col), flag(era), new_row);
    let replacement = format!("{head}:{new_end}");
    let named = if qualifier.is_empty() {
        let plain = sheet_name.chars().all(|c| c.is_alphanumeric() || c == '_');
        if plain {
            format!("{sheet_name}!")
        } else {
            format!("'{}'!", sheet_name.replace('\'', "''"))
        }
    } else {
        qualifier.to_string()
    };
    let added_start = format!(
        "{}{}{}{}",
        flag(sca),
        letters(first_col),
        flag(sra),
        first_row
    );
    let added = if added_start == new_end {
        format!("{named}{new_end}")
    } else {
        format!("{named}{added_start}:{new_end}")
    };
    (replacement, added)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet, Workbook};
    use crate::rules::LinterContext;
    use crate::rules::collect_violations;
    use std::collections::HashMap;
    use std::sync::Arc;

    fn make_number_cell(row: u32, col: u32, n: f64) -> ((u32, u32), Cell) {
        (
            (row, col),
            Cell {
                value: CellValue::Number(n),
                formula: None,
                is_array: false,
                num_fmt: None,
            },
        )
    }

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

    fn make_formula_cell(row: u32, col: u32, formula: &str) -> ((u32, u32), Cell) {
        (
            (row, col),
            Cell {
                value: CellValue::Empty,
                formula: Some(Box::from(formula)),
                is_array: false,
                num_fmt: None,
            },
        )
    }

    fn make_sheet_with_index(index: u16, name: &str, cells: Vec<((u32, u32), Cell)>) -> Sheet {
        let mut cell_map = HashMap::new();
        for (at, cell) in cells {
            cell_map.insert(at, cell);
        }
        Sheet {
            name: name.to_string(),
            sheet_index: index,
            cells: cell_map.into(),
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

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
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

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
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

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(
            violations.len(),
            0,
            "Should suppress when adjacent cell is text but range is numeric"
        );
    }

    fn lint_formula(cells: Vec<((u32, u32), Cell)>, formula_loc: (u32, u32)) -> Vec<Violation> {
        let sheet = make_sheet_with_index(0, "Sheet1", cells);
        let workbook = make_workbook(vec![sheet]);
        let mut rule = LongerRefExpectedRule::new(&LinterConfig::default());
        let mut ctx = LinterContext::default();
        ctx.name_to_index.insert("Sheet1".to_string(), 0);
        ctx.cell_dependencies
            .insert((0, formula_loc.0, formula_loc.1), Vec::new());
        collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out))
    }

    #[test]
    fn test_type_match_uses_last_range_cell() {
        // A1:A2 numbers and A3:A4 text tie by count; A4 (text) decides.
        let mut cells = vec![
            make_number_cell(0, 0, 1.0),
            make_number_cell(1, 0, 2.0),
            make_text_cell(2, 0, "a"),
            make_text_cell(3, 0, "b"),
            make_formula_cell(0, 1, "COUNTA(A1:A4)"),
        ];
        cells.push(make_text_cell(4, 0, "c"));
        assert_eq!(lint_formula(cells.clone(), (0, 1)).len(), 1);

        cells.pop();
        cells.push(make_number_cell(4, 0, 3.0));
        assert!(lint_formula(cells, (0, 1)).is_empty());
    }

    #[test]
    fn test_type_match_requires_every_boundary_pair() {
        // A1:B2 numbers; row 3 is compared column by column.
        let mut cells = vec![
            make_number_cell(0, 0, 1.0),
            make_number_cell(0, 1, 1.0),
            make_number_cell(1, 0, 1.0),
            make_number_cell(1, 1, 1.0),
            make_number_cell(2, 0, 1.0),
            make_formula_cell(0, 3, "SUM(A1:B2)"),
        ];
        assert!(
            lint_formula(cells.clone(), (0, 3)).is_empty(),
            "B3 empty next to numeric B2 is a different type"
        );

        cells.push(make_number_cell(2, 1, 1.0));
        assert_eq!(lint_formula(cells, (0, 3)).len(), 1);
    }

    #[test]
    fn test_empty_last_cell_differs_from_data() {
        // A3 empty, A4 number → different types → no flag.
        let cells = vec![
            make_number_cell(0, 0, 1.0),
            make_number_cell(1, 0, 2.0),
            make_number_cell(3, 0, 4.0),
            make_formula_cell(0, 1, "SUM(A1:A3)"),
        ];
        assert!(lint_formula(cells, (0, 1)).is_empty());
    }

    #[test]
    fn test_uncached_formula_matches_any_type() {
        // A3 holds a formula without a cached value.
        let base = vec![
            make_number_cell(0, 0, 1.0),
            make_number_cell(1, 0, 2.0),
            make_formula_cell(0, 1, "SUM(A1:A3)"),
        ];

        let mut last_uncached = base.clone();
        last_uncached.push(make_formula_cell(2, 0, "B2*2"));
        last_uncached.push(make_formula_cell(3, 0, "B3*2"));
        assert_eq!(lint_formula(last_uncached, (0, 1)).len(), 1);

        // A formula next to the range is not data, cached or not.
        let mut adjacent_uncached = base;
        adjacent_uncached.push(make_number_cell(2, 0, 3.0));
        adjacent_uncached.push(make_formula_cell(3, 0, "A3*2"));
        assert!(lint_formula(adjacent_uncached, (0, 1)).is_empty());
    }

    #[test]
    fn test_every_range_in_formula_is_evaluated() {
        // A1:A3 is followed by a text header (suppressed); C1:C3 by a number.
        let mut cells = Vec::new();
        for row in 0..3 {
            cells.push(make_number_cell(row, 0, 1.0));
            cells.push(make_number_cell(row, 2, 1.0));
        }
        cells.push(make_text_cell(3, 0, "Total"));
        cells.push(make_number_cell(3, 2, 1.0));
        cells.push(make_formula_cell(0, 4, "SUM(A1:A3)+SUM(C1:C3)"));

        let violations = lint_formula(cells.clone(), (0, 4));
        assert_eq!(violations.len(), 1);
        assert_eq!(
            violations[0].data::<LongerRefData>().unwrap().range_str,
            "C1:C3"
        );

        cells.retain(|(at, _)| *at != (3, 0));
        cells.push(make_number_cell(3, 0, 1.0));
        let ranges: Vec<_> = lint_formula(cells, (0, 4))
            .iter()
            .map(|v| v.data::<LongerRefData>().unwrap().range_str.clone())
            .collect();
        assert_eq!(ranges, ["A1:A3", "C1:C3"]);
    }

    #[test]
    fn test_repeated_range_reported_once() {
        let mut cells: Vec<_> = (0..4).map(|row| make_number_cell(row, 0, 1.0)).collect();
        cells.push(make_formula_cell(0, 2, "SUM(A1:A3)/COUNT($A$1:$A$3)"));
        assert_eq!(lint_formula(cells, (0, 2)).len(), 1);
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

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(violations.len(), 1);
    }

    #[test]
    fn test_no_formula() {
        let sheet = make_sheet_with_index(0, "Sheet1", vec![make_number_cell(0, 0, 42.0)]);
        let workbook = make_workbook(vec![sheet]);

        let config = LinterConfig::default();
        let mut rule = LongerRefExpectedRule::new(&config);
        let ctx = &mut LinterContext::default();

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, ctx, out));
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

        let violations = collect_violations(|out| rule.on_workbook_end(&workbook, &mut ctx, out));
        assert_eq!(
            violations.len(),
            0,
            "min_adjacent_data=2 but only 1 adjacent → no flag"
        );
    }

    #[test]
    fn a_block_of_formulas_is_extended_by_the_same_formula_pattern() {
        // B1:B3 = A*2 per row; the range sums B1:B3 and B4 repeats the pattern.
        let block = |row4: &str| {
            let mut cells: Vec<_> = (0..4).map(|r| make_number_cell(r, 0, 1.0)).collect();
            for r in 0..3 {
                cells.push(make_formula_cell(r, 1, &format!("A{}*2", r + 1)));
            }
            cells.push(make_formula_cell(3, 1, row4));
            cells.push(make_formula_cell(0, 3, "SUM(B1:B3)"));
            lint_formula(cells, (0, 3)).len()
        };
        assert_eq!(block("A4*2"), 1, "same pattern continues the block");
        assert_eq!(block("SUM(B1:B3)"), 0, "a different formula is a result");
    }

    #[test]
    fn constants_and_formulas_do_not_extend_each_other() {
        // Constants summed, a formula below: a total, not data the range lacks.
        let mut cells: Vec<_> = (0..3).map(|r| make_number_cell(r, 1, 1.0)).collect();
        cells.push(make_formula_cell(3, 1, "B3*2"));
        cells.push(make_formula_cell(0, 3, "SUM(B1:B3)"));
        assert!(lint_formula(cells, (0, 3)).is_empty());
        // Formulas summed, constants below: inputs, not a continuation of the calculation.
        let mut cells: Vec<_> = (0..3)
            .map(|r| make_formula_cell(r, 1, &format!("A{}*2", r + 1)))
            .collect();
        cells.push(make_number_cell(3, 1, 5.0));
        cells.push(make_formula_cell(0, 3, "SUM(B1:B3)"));
        assert!(lint_formula(cells, (0, 3)).is_empty());
    }

    #[test]
    fn suggestion_keeps_the_format_of_the_range() {
        let s = |d: &str, sheet: &str, vertical: bool, n: usize| suggestion(d, sheet, vertical, n);
        assert_eq!(
            s("INGRESOS!$H$61:$Q$61", "INGRESOS", false, 10),
            (
                "INGRESOS!$H$61:$AA$61".to_string(),
                "INGRESOS!$R$61:$AA$61".to_string()
            )
        );
        assert_eq!(
            s("$L$253:CQ$253", "Optimizacion", false, 3),
            (
                "$L$253:CT$253".to_string(),
                "Optimizacion!$CR$253:CT$253".to_string()
            )
        );
        assert_eq!(
            s("$C$585:$FK$616", "INGRESOS", false, 40),
            (
                "$C$585:$GY$616".to_string(),
                "INGRESOS!$FL$585:$GY$616".to_string()
            )
        );
        assert_eq!(
            s("A1:B5", "S", true, 2),
            ("A1:B7".to_string(), "S!A6:B7".to_string())
        );
        assert_eq!(
            s("A1:A5", "Mi hoja", true, 3),
            ("A1:A8".to_string(), "'Mi hoja'!A6:A8".to_string())
        );
        assert_eq!(
            s("'Mi hoja'!B$2:B$4", "Mi hoja", true, 1),
            ("'Mi hoja'!B$2:B$5".to_string(), "'Mi hoja'!B$5".to_string())
        );
    }

    #[test]
    fn the_message_gives_the_whole_block_and_the_replacement() {
        // A1:A3 summed; the block of numbers goes on to A7.
        let mut cells: Vec<_> = (0..7).map(|r| make_number_cell(r, 0, 1.0)).collect();
        cells.push(make_formula_cell(0, 2, "SUM(A1:A3)"));
        let violations = lint_formula(cells, (0, 2));
        assert_eq!(violations.len(), 1);
        let data = violations[0].data::<LongerRefData>().unwrap();
        assert_eq!(data.lines, 4);
        assert_eq!(data.replacement, "A1:A7");
        assert_eq!(data.added, "Sheet1!A4:A7");
    }
}

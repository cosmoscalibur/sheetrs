//! Common helper functions shared across rule implementations.
//!
//! Provides cell grouping, bounding-box utilities, formula-parsing helpers,
//! and a shared cell-reference regex used by multiple rules.

use regex::Regex;
use std::collections::{HashMap, HashSet, VecDeque};
use std::sync::LazyLock;

/// Compiled cell-reference pattern shared across rules.
///
/// Matches cell references like `A1`, `$A$1`, `Sheet1!A1`, `'Sheet Name'!A1`,
/// `A1:B2`.
///
/// Capture groups:
/// - 1: full sheet qualifier (including `!`)
/// - 2: quoted sheet name (inside single quotes)
/// - 3: unquoted sheet name
/// - 4: start column letters
/// - 5: start row number
/// - 6: end column letters (optional, range only)
/// - 7: end row number (optional, range only)
pub static CELL_REF_PATTERN: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(
        r"(?:('([^']+)'|([A-Za-z0-9_\.]+))!)?\$?([A-Za-z]+)\$?([0-9]+)(?::\$?([A-Za-z]+)\$?([0-9]+))?",
    )
    .expect("cell reference regex must compile")
});

/// Parse a row string and column string into zero-based `(row, col)` coordinates.
///
/// The row string is a 1-based decimal number. The column string is a sequence
/// of ASCII letters representing a base-26 column index (A=1, Z=26, AA=27, …).
///
/// Returns `None` unless the pair is a cell of a sheet, up to column `XFD` and
/// row 1,048,576: a token such as `A0`, `ABCD1` or `TOTALSALES2024` has the
/// shape of a reference and is a name.
pub fn parse_cell_coords(row_str: &str, col_str: &str) -> Option<(u32, u32)> {
    let row = row_str
        .parse::<u32>()
        .ok()
        .filter(|row| (1..=sheetrs_formula::addr::MAX_ROWS).contains(row))?;
    let col = sheetrs_formula::addr::col_index(col_str)?;
    Some((row - 1, col))
}

/// Compute the bounding box for a set of cells.
///
/// Returns `(min_row, min_col, max_row, max_col)`. Falls back to `(0,0,0,0)`
/// when `cells` is empty.
pub fn bounding_box(cells: &[(u32, u32)]) -> (u32, u32, u32, u32) {
    let min_row = cells.iter().map(|(r, _)| *r).min().unwrap_or(0);
    let min_col = cells.iter().map(|(_, c)| *c).min().unwrap_or(0);
    let max_row = cells.iter().map(|(r, _)| *r).max().unwrap_or(0);
    let max_col = cells.iter().map(|(_, c)| *c).max().unwrap_or(0);
    (min_row, min_col, max_row, max_col)
}

/// Group cells into contiguous (BFS 4-connected) ranges.
///
/// Returns a `Vec` where each element is a group of cells that form a
/// contiguous block via row/column adjacency.
pub fn find_contiguous_ranges(cells: &[(u32, u32)]) -> Vec<Vec<(u32, u32)>> {
    let cell_set: HashSet<(u32, u32)> = cells.iter().copied().collect();
    let mut visited: HashSet<(u32, u32)> = HashSet::new();
    let mut ranges: Vec<Vec<(u32, u32)>> = Vec::new();

    for &cell in cells {
        if visited.contains(&cell) {
            continue;
        }

        let mut range = Vec::new();
        let mut queue = VecDeque::new();
        queue.push_back(cell);
        visited.insert(cell);

        while let Some((row, col)) = queue.pop_front() {
            range.push((row, col));

            let neighbors = [
                (row.wrapping_sub(1), col),
                (row + 1, col),
                (row, col.wrapping_sub(1)),
                (row, col + 1),
            ];

            for neighbor in neighbors {
                if cell_set.contains(&neighbor) && !visited.contains(&neighbor) {
                    visited.insert(neighbor);
                    queue.push_back(neighbor);
                }
            }
        }

        ranges.push(range);
    }

    ranges
}

/// Format a group of cells as a single range string (e.g., `A1:C5`).
///
/// Returns the single cell reference for single-cell groups,
/// or a `start:end` bounding-box range for multi-cell groups.
pub fn format_single_range(cells: &[(u32, u32)]) -> String {
    use crate::violation::CellReference;

    if cells.is_empty() {
        return String::new();
    }

    if cells.len() == 1 {
        return CellReference::new(cells[0].0, cells[0].1).to_string();
    }

    let (min_row, min_col, max_row, max_col) = bounding_box(cells);
    let start = CellReference::new(min_row, min_col);
    let end = CellReference::new(max_row, max_col);
    format!("{start}:{end}")
}

/// Check if a position in a formula string is inside a double-quoted string literal.
///
/// Scans from the start of `formula` up to `pos`, toggling an in-string flag
/// on each `"` character. Returns `true` if `pos` falls inside a literal.
pub fn is_inside_string(formula: &str, pos: usize) -> bool {
    let mut in_string = false;
    for (i, ch) in formula.char_indices() {
        if i >= pos {
            break;
        }
        if ch == '"' {
            in_string = !in_string;
        }
    }
    in_string
}

/// Extract function arguments by balancing parentheses from `paren_pos`.
///
/// `paren_pos` is the index of the opening `(` in `formula`.
/// Returns `None` if the parentheses are unbalanced.
/// Returns `Some(Vec<String>)` with each top-level argument as a string.
///
/// Nested parentheses (e.g., `SUM(IF(cond,A1,A2),B1)`) are preserved inside
/// the argument string — only top-level commas split arguments.
pub fn extract_args(formula: &str, paren_pos: usize) -> Option<Vec<String>> {
    let bytes = formula.as_bytes();
    if paren_pos >= bytes.len() || bytes[paren_pos] != b'(' {
        return None;
    }

    let mut depth = 0;
    let mut args = Vec::new();
    let mut current_arg = String::new();
    let mut in_string = false;

    // SAFETY: We only match ASCII chars (, ) , " whose byte values (0x22, 0x28–0x2C)
    // cannot appear inside multi-byte UTF-8 sequences (all continuation bytes are >= 0x80).
    for &b in &bytes[paren_pos..] {
        match b {
            b'"' => {
                in_string = !in_string;
                current_arg.push(b as char);
            }
            b'(' if !in_string => {
                depth += 1;
                if depth > 1 {
                    current_arg.push('(');
                }
            }
            b')' if !in_string => {
                depth -= 1;
                if depth == 0 {
                    // End of function call
                    if !current_arg.is_empty() || !args.is_empty() {
                        args.push(current_arg);
                    }
                    return Some(args);
                }
                current_arg.push(')');
            }
            b',' if !in_string && depth == 1 => {
                args.push(current_arg);
                current_arg = String::new();
            }
            _ => {
                if depth >= 1 {
                    current_arg.push(b as char);
                }
            }
        }
    }

    None // Unbalanced parentheses
}

/// Parse a range reference string (e.g., `A1:B10`, `$A$1:$B$10`, or `A1`)
/// into 0-based coordinates `(start_row, start_col, end_row, end_col)`.
///
/// For single-cell references, start and end coordinates are equal.
///
/// Returns `None` unless every end is a cell of a sheet (up to column `XFD` and
/// row 1,048,576), so a nested call such as `ABS(A1)` or a whole-column `A:A`
/// is not a range.
pub fn parse_formula_range(range_str: &str) -> Option<(u32, u32, u32, u32)> {
    let cell = |s: &str| sheetrs_formula::addr::parse_cell(s).map(|(row, col, ..)| (row, col));
    match range_str.split_once(':') {
        Some((start, end)) => {
            let (sr, sc) = cell(start)?;
            let (er, ec) = cell(end)?;
            Some((sr, sc, er, ec))
        }
        None => {
            let (r, c) = cell(range_str)?;
            Some((r, c, r, c))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The limits of LibreOffice 26.8: column `XFD` and row 1,048,576. A token beyond them or on
    /// row 0 is a name, as `=A0` and `=XFE1` are `#NAME?` there.
    #[test]
    fn test_parse_cell_coords_accepts_only_cells_of_a_sheet() {
        assert_eq!(parse_cell_coords("1", "A"), Some((0, 0)));
        assert_eq!(
            parse_cell_coords("1048576", "XFD"),
            Some((1_048_575, 16_383))
        );
        assert_eq!(parse_cell_coords("0", "A"), None);
        assert_eq!(parse_cell_coords("1048577", "A"), None);
        assert_eq!(parse_cell_coords("1", "XFE"), None);
        assert_eq!(parse_cell_coords("1", "ABCD"), None);
        assert_eq!(parse_cell_coords("2024", "TOTALSALES"), None);
        assert_eq!(parse_cell_coords("99999999999", "A"), None);
    }

    #[test]
    fn test_parse_formula_range_rejects_what_is_not_a_cell() {
        assert_eq!(parse_formula_range("A1"), Some((0, 0, 0, 0)));
        assert_eq!(parse_formula_range("$A$1:$B$10"), Some((0, 0, 9, 1)));
        // A nested call and a whole column.
        assert_eq!(
            parse_formula_range("ABS(MIN(MAX(ROUND(AVERAGE(A1),2),0),100),0)"),
            None
        );
        assert_eq!(parse_formula_range("A:A"), None);
        assert_eq!(parse_formula_range("A0"), None);
    }

    #[test]
    fn test_bounding_box_single_cell() {
        assert_eq!(bounding_box(&[(3, 5)]), (3, 5, 3, 5));
    }

    #[test]
    fn test_bounding_box_multiple_cells() {
        assert_eq!(bounding_box(&[(0, 0), (5, 3), (2, 7)]), (0, 0, 5, 7));
    }

    #[test]
    fn test_bounding_box_empty() {
        assert_eq!(bounding_box(&[]), (0, 0, 0, 0));
    }

    #[test]
    fn test_contiguous_single_group() {
        let cells = vec![(0, 0), (0, 1), (1, 0)];
        let ranges = find_contiguous_ranges(&cells);
        assert_eq!(ranges.len(), 1);
        assert_eq!(ranges[0].len(), 3);
    }

    #[test]
    fn test_contiguous_two_groups() {
        let cells = vec![(0, 0), (0, 1), (5, 5)];
        let ranges = find_contiguous_ranges(&cells);
        assert_eq!(ranges.len(), 2);
    }

    #[test]
    fn test_contiguous_empty() {
        let ranges = find_contiguous_ranges(&[]);
        assert!(ranges.is_empty());
    }

    // --- is_inside_string ---

    #[test]
    fn test_not_inside_string() {
        assert!(!is_inside_string("SUM(A1)", 0));
    }

    #[test]
    fn test_inside_string() {
        // Position 5 is inside the quoted portion: "hello"
        assert!(is_inside_string("=\"hello\"", 5));
    }

    #[test]
    fn test_after_string() {
        // Position 8 is after the closing quote
        assert!(!is_inside_string("=\"hello\"+A1", 8));
    }

    // --- extract_args ---

    #[test]
    fn test_extract_args_simple() {
        let args = extract_args("SUM(A1,B1,C1)", 3).unwrap();
        assert_eq!(args, vec!["A1", "B1", "C1"]);
    }

    #[test]
    fn test_extract_args_nested() {
        let args = extract_args("SUM(IF(cond,A1,A2),B1)", 3).unwrap();
        assert_eq!(args.len(), 2);
        assert_eq!(args[0], "IF(cond,A1,A2)");
        assert_eq!(args[1], "B1");
    }

    #[test]
    fn test_extract_args_single() {
        let args = extract_args("SUM(A1:A10)", 3).unwrap();
        assert_eq!(args, vec!["A1:A10"]);
    }

    #[test]
    fn test_extract_args_empty_parens() {
        let args = extract_args("NOW()", 3).unwrap();
        assert!(args.is_empty());
    }

    #[test]
    fn test_extract_args_unbalanced() {
        assert!(extract_args("SUM(A1,B1", 3).is_none());
    }

    // --- parse_formula_range ---

    #[test]
    fn test_parse_range() {
        assert_eq!(parse_formula_range("A1:B10"), Some((0, 0, 9, 1)));
    }

    #[test]
    fn test_parse_range_absolute() {
        assert_eq!(parse_formula_range("$A$1:$B$10"), Some((0, 0, 9, 1)));
    }

    #[test]
    fn test_parse_single_cell() {
        assert_eq!(parse_formula_range("C5"), Some((4, 2, 4, 2)));
    }

    #[test]
    fn test_parse_whole_column() {
        // Whole-column refs like A:A have no row digits → parse_cell_ref returns None
        assert_eq!(parse_formula_range("A:A"), None);
    }

    // --- parse_cell_coords ---

    #[test]
    fn test_parse_cell_coords_a1() {
        assert_eq!(parse_cell_coords("1", "A"), Some((0, 0)));
    }

    #[test]
    fn test_parse_cell_coords_z26() {
        assert_eq!(parse_cell_coords("26", "Z"), Some((25, 25)));
    }

    #[test]
    fn test_parse_cell_coords_aa() {
        assert_eq!(parse_cell_coords("1", "AA"), Some((0, 26)));
    }

    #[test]
    fn test_parse_cell_coords_invalid_row() {
        assert_eq!(parse_cell_coords("abc", "A"), None);
    }

    #[test]
    fn test_parse_cell_coords_empty_col() {
        assert_eq!(parse_cell_coords("1", ""), None);
    }

    // --- CELL_REF_PATTERN ---

    #[test]
    fn test_regex_simple_ref() {
        let caps = CELL_REF_PATTERN.captures("A1").unwrap();
        assert_eq!(caps.get(4).unwrap().as_str(), "A");
        assert_eq!(caps.get(5).unwrap().as_str(), "1");
        assert!(caps.get(6).is_none());
    }

    #[test]
    fn test_regex_range_ref() {
        let caps = CELL_REF_PATTERN.captures("B2:D10").unwrap();
        assert_eq!(caps.get(4).unwrap().as_str(), "B");
        assert_eq!(caps.get(5).unwrap().as_str(), "2");
        assert_eq!(caps.get(6).unwrap().as_str(), "D");
        assert_eq!(caps.get(7).unwrap().as_str(), "10");
    }

    #[test]
    fn test_regex_sheet_qualified() {
        let caps = CELL_REF_PATTERN.captures("Sheet1!A1").unwrap();
        assert_eq!(caps.get(3).unwrap().as_str(), "Sheet1");
        assert_eq!(caps.get(4).unwrap().as_str(), "A");
    }

    #[test]
    fn test_regex_quoted_sheet() {
        let caps = CELL_REF_PATTERN.captures("'My Sheet'!A1").unwrap();
        assert_eq!(caps.get(2).unwrap().as_str(), "My Sheet");
    }
}

/// What a formula links to: the sheets whose cells, ranges, whole columns or whole rows it names
/// (a reference without a sheet is the formula's own sheet) and the defined names it uses
/// (uppercase).
#[derive(Debug, Default, PartialEq, Eq)]
pub struct FormulaLinks {
    /// Sheet indexes, in order of appearance, repeated as in the formula.
    pub sheets: Vec<u16>,
    /// Defined names, uppercase.
    pub names: Vec<String>,
}

/// The links of a formula, from its parsed form (a reference-like text inside a string literal
/// is not one). `None` when the formula cannot be parsed.
pub fn formula_links(
    formula: &str,
    current_sheet: u16,
    sheet_indexes: &HashMap<String, u16>,
) -> Option<FormulaLinks> {
    use sheetrs_formula::ast::Expr;
    fn walk(e: &Expr, current: u16, indexes: &HashMap<String, u16>, out: &mut FormulaLinks) {
        match e {
            Expr::Ref(r) => match &r.sheet {
                None => out.sheets.push(current),
                Some(name) => {
                    if let Some(&i) = indexes.get(name.as_str()) {
                        out.sheets.push(i);
                    }
                }
            },
            Expr::Name(n) => out.names.push(n.to_uppercase()),
            Expr::Call(c) => c
                .args
                .iter()
                .flatten()
                .for_each(|a| walk(a, current, indexes, out)),
            Expr::Array(rows) => rows
                .iter()
                .flatten()
                .for_each(|x| walk(x, current, indexes, out)),
            Expr::Unary(_, x) | Expr::Percent(x) | Expr::Paren(x) => walk(x, current, indexes, out),
            Expr::Binary(_, a, b) => {
                walk(a, current, indexes, out);
                walk(b, current, indexes, out);
            }
            Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Error(_) => {}
        }
    }
    let ast = sheetrs_formula::parse(formula.strip_prefix('=').unwrap_or(formula)).ok()?;
    let mut out = FormulaLinks::default();
    walk(&ast, current_sheet, sheet_indexes, &mut out);
    Some(out)
}

/// The formula with the text of its string literals (`"..."`) and quoted sheet names (`'...'`)
/// replaced by spaces, so that what is searched for in the rest cannot be found inside them. Quotes
/// are doubled to escape them in both (`"say ""hi"""`).
pub fn mask_quoted(formula: &str) -> std::borrow::Cow<'_, str> {
    if !formula.contains(['"', '\'']) {
        return std::borrow::Cow::Borrowed(formula);
    }
    let mut out = String::with_capacity(formula.len());
    let mut quote: Option<char> = None;
    let mut chars = formula.chars().peekable();
    while let Some(c) = chars.next() {
        match quote {
            None => {
                if c == '"' || c == '\'' {
                    quote = Some(c);
                    out.push(' ');
                } else {
                    out.push(c);
                }
            }
            Some(q) => {
                out.push(' ');
                if c == q {
                    if chars.peek() == Some(&q) {
                        chars.next();
                        out.push(' ');
                    } else {
                        quote = None;
                    }
                }
            }
        }
    }
    std::borrow::Cow::Owned(out)
}

/// The prefixed (newer than the file format) function calls of a formula, outside its quoted text.
#[derive(Debug, Default, PartialEq, Eq)]
pub struct FutureCalls {
    /// The formula calls some function written with the `_xlfn.` or `_xlws.` prefix.
    pub any: bool,
    /// One of them is `_xlfn._LONGTEXT(`, which is not an Excel function: LibreOffice 26.2 wrote a
    /// text literal of more than 255 characters as pieces joined by it.
    pub longtext: bool,
}

/// The prefixed function calls of `formula`.
pub fn future_calls(formula: &str) -> FutureCalls {
    if !formula.contains(['_', '.']) {
        return FutureCalls::default();
    }
    let masked = mask_quoted(formula).to_ascii_uppercase();
    FutureCalls {
        any: masked.contains("_XLFN.") || masked.contains("_XLWS."),
        longtext: masked.contains("_XLFN._LONGTEXT("),
    }
}

/// Whether `name(` (uppercase) appears as a function call in an uppercase formula whose quoted
/// text has been masked ([`mask_quoted`]): not as the end of a longer name (`SUMIF(` is not `IF(`)
/// and not as the end of a prefixed one (`_XLFN.IFS(`) or an external reference.
pub fn has_call(masked_upper: &str, name: &str) -> bool {
    let call = format!("{name}(");
    masked_upper.match_indices(&call).any(|(i, _)| {
        !masked_upper[..i]
            .chars()
            .next_back()
            .is_some_and(|p| p.is_ascii_alphanumeric() || p == '_' || p == '.')
    })
}

/// Whether a formula names a whole column (`A:A`) and whether it names a whole row (`1:1`), from
/// its parsed form. `None` when the formula cannot be parsed.
pub fn whole_line_refs(formula: &str) -> Option<(bool, bool)> {
    use sheetrs_formula::ast::{Expr, RefArea};
    fn walk(e: &Expr, found: &mut (bool, bool)) {
        match e {
            Expr::Ref(r) => match r.area {
                RefArea::Cols(..) => found.0 = true,
                RefArea::Rows(..) => found.1 = true,
                _ => {}
            },
            Expr::Call(c) => c.args.iter().flatten().for_each(|a| walk(a, found)),
            Expr::Array(rows) => rows.iter().flatten().for_each(|x| walk(x, found)),
            Expr::Unary(_, x) | Expr::Percent(x) | Expr::Paren(x) => walk(x, found),
            Expr::Binary(_, a, b) => {
                walk(a, found);
                walk(b, found);
            }
            Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Error(_) | Expr::Name(_) => {}
        }
    }
    let ast = sheetrs_formula::parse(formula.strip_prefix('=').unwrap_or(formula)).ok()?;
    let mut found = (false, false);
    walk(&ast, &mut found);
    Some(found)
}

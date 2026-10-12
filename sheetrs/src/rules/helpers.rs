//! Common helper functions shared across rule implementations.
//!
//! Provides cell grouping, bounding-box utilities, formula-parsing helpers,
//! and a shared cell-reference regex used by multiple rules.

use crate::reader::parser_utils::column_letters_to_index;
use regex::Regex;
use std::collections::{HashSet, VecDeque};
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
/// Returns `None` if the row cannot be parsed or the column string is empty.
pub fn parse_cell_coords(row_str: &str, col_str: &str) -> Option<(u32, u32)> {
    let row = row_str.parse::<u32>().ok()?;
    let col = column_letters_to_index(col_str)?;
    Some((row.saturating_sub(1), col))
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
/// Strips `$` signs before delegating to `parse_cell_ref`. For single-cell
/// references, start and end coordinates are equal.
///
/// Returns `None` if the reference cannot be parsed (e.g., whole-column `A:A`).
pub fn parse_formula_range(range_str: &str) -> Option<(u32, u32, u32, u32)> {
    use crate::reader::parser_utils::parse_cell_ref;

    // Strip $ signs for absolute reference handling
    let cleaned: String = range_str.chars().filter(|c| *c != '$').collect();

    // Validate each part looks like a cell ref before passing to parse_cell_ref.
    // Excel columns have at most 3 letters (max: XFD). Reject anything with
    // more alpha chars to prevent arithmetic overflow in parse_cell_ref.
    let validate_ref = |s: &str| -> bool {
        let alpha_count = s.chars().take_while(|c| c.is_ascii_alphabetic()).count();
        alpha_count > 0 && alpha_count <= 3 && s.len() > alpha_count
    };

    if let Some((start, end)) = cleaned.split_once(':') {
        if !validate_ref(start) || !validate_ref(end) {
            return None;
        }
        let (sr, sc) = parse_cell_ref(start)?;
        let (er, ec) = parse_cell_ref(end)?;
        Some((sr, sc, er, ec))
    } else {
        // Single cell reference
        if !validate_ref(&cleaned) {
            return None;
        }
        let (r, c) = parse_cell_ref(&cleaned)?;
        Some((r, c, r, c))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

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

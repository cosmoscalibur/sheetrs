//! `^` and `$` in the regular expression of a search (`find`): they anchor the text of the cell, in
//! an XLSX and in an ODS. `libreoffice_find.rs` compares many such searches with what `LibreOffice`
//! finds in the same file; these say the rules, on `tests/assets/anchors.*` (`T!A1:A21`), with the
//! expectations that `LibreOffice` gave (`make_find_oracle.py`).
//!
//! Rows: 1 `a`, 2 `ab`, 3 `ba`, 4 `b`, 5 `A`, 6 `aa`, 7 `a b`, 8 `Code`, 9 `code`, 10 `Code x`,
//! 11 `Resumen Contadia`, 12 `x⏎a`, 13 `a⏎x`, 14 `a⏎x⏎b`, 15 `x⏎a⏎b`, 16 `ab⏎cd`, 17 `a⏎`, 18 `⏎a`,
//! 19 `a⏎⏎`, 20 `a` and U+2028, 21 U+2028 and `a`.

use sheetrs_io::Book;
use sheetrs_protocol::{Find, Mode, Query};
use sheetrs_query::BookQuery;
use std::path::Path;

fn search(extension: &str, text: &str, entire_cells: bool, match_case: bool) -> Vec<u32> {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(format!("anchors.{extension}"));
    let find = Find {
        mode: Mode::Regex,
        entire_cells,
        match_case,
        ..Find::literal(text)
    };
    let query = Query {
        find: Some(find),
        ..Query::default()
    };
    let mut rows: Vec<u32> = Book::from_file(path)
        .unwrap()
        .query(&query)
        .unwrap()
        .cells
        .iter()
        .map(|c| c.cell()[1..].parse().unwrap())
        .collect();
    rows.sort_unstable();
    rows
}

/// The rows found in the XLSX, which are those found in the ODS.
fn rows(text: &str, entire_cells: bool, match_case: bool) -> Vec<u32> {
    let xlsx = search("xlsx", text, entire_cells, match_case);
    assert_eq!(
        xlsx,
        search("ods", text, entire_cells, match_case),
        "{text}"
    );
    xlsx
}

#[test]
fn the_names_of_a_header_row_are_found_by_exact_text() {
    // `Code` and `Resumen Contadia`; with match case, not `code`; never `Code x`.
    for entire_cells in [false, true] {
        assert_eq!(
            rows("^(Code|Resumen Contadia)$", entire_cells, true),
            [8, 11]
        );
        assert_eq!(
            rows("^(Code|Resumen Contadia)$", entire_cells, false),
            [8, 9, 11]
        );
    }
}

#[test]
fn an_anchor_is_the_start_and_the_end_of_the_text_not_of_its_lines() {
    // `x⏎a` (12) and `x⏎a⏎b` (15) start with `x`; `a⏎x` (13) and `Code x` (10) end with it.
    assert_eq!(rows("^x", false, false), [12, 15]);
    assert_eq!(rows("x$", false, false), [10, 13]);
    assert_eq!(rows("^x$", false, false), Vec::<u32>::new());
    // `x⏎b`-like second lines do not start the text: only `ba` (3) and `b` (4) begin with `b`.
    assert_eq!(rows("^b", false, false), [3, 4]);
}

#[test]
fn a_dollar_matches_before_the_line_break_that_ends_the_text_only() {
    // `a⏎` (17) and `a` with U+2028 (20) end in one terminator; `a⏎⏎` (19) in two.
    assert_eq!(rows("^a$", false, false), [1, 5, 17, 20]);
    assert_eq!(
        rows("a$", false, false),
        [1, 3, 5, 6, 11, 12, 17, 18, 20, 21]
    );
    // The whole text is the cell, so a trailing line break is not part of the match.
    assert_eq!(rows("^a$", true, false), [1, 5]);
}

#[test]
fn an_anchor_inside_a_pattern_is_where_the_pattern_puts_it() {
    assert_eq!(rows("a^b", false, false), Vec::<u32>::new());
    assert_eq!(rows("a$b", false, false), Vec::<u32>::new());
    assert_eq!(
        rows("(^a|b$)", false, false),
        [1, 2, 4, 5, 6, 7, 13, 14, 15, 16, 17, 19, 20]
    );
    // The first match is the whole text for `entire_cells`, as without anchors.
    assert_eq!(rows("(^a|b$)", true, false), [1, 4, 5]);
}

#[test]
fn a_dot_still_leaves_out_a_line_break() {
    // `^.` is satisfied by the first character of `x⏎a` (12) but not by the line break of `⏎a` (18).
    let first = rows("^.", false, false);
    assert!(first.contains(&12) && !first.contains(&18));
    assert_eq!(rows("a.$", false, false), [2, 6]);
}

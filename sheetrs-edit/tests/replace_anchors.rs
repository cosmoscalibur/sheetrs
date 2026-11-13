//! `^` and `$` in the regular expression of a replacement, and the patterns that can match the
//! empty text, against `LibreOffice` itself.
//!
//! `tests/assets/replace_anchors/` has a workbook made with `LibreOffice` 26.8 in both formats
//! (`input.*`, `T!A1:A..`, one text a row, among them texts of several lines) and `oracle.json`,
//! the text of each cell after `LibreOffice` ran `Sheet.replaceAll` with a regular expression, for
//! each search, `match_case` and `entire_cells` (`make_anchors_oracle.py`). The editor runs the
//! same replacements and every cell must be what `LibreOffice` left.
//!
//! What it says, in short: a match of no characters is not replaced, unless it is at the very end of
//! the text; the search that follows a hit begins where the hit ended, and none begins at the end of
//! the text; `$` is the end of the text or the place before the line terminator that ends it, `^`
//! the start of the text (not of a line).

use serde_json::Value as Json;
use sheetrs_edit::BookEdit;
use sheetrs_io::Book;
use sheetrs_io::reader::CellValue;
use sheetrs_protocol::{Find, Mode, Replace};
use std::path::{Path, PathBuf};
use std::sync::LazyLock;

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/replace_anchors")
        .join(name)
}

static ORACLE: LazyLock<Json> = LazyLock::new(|| {
    serde_json::from_str(&std::fs::read_to_string(asset("oracle.json")).unwrap()).unwrap()
});

fn texts() -> Vec<String> {
    ORACLE["texts"]
        .as_array()
        .unwrap()
        .iter()
        .map(|t| t.as_str().unwrap().to_string())
        .collect()
}

/// What the editor leaves in each cell of `T!A:A` after replacing `find` by `with`.
fn replaced(
    extension: &str,
    find: &str,
    with: &str,
    match_case: bool,
    entire_cells: bool,
) -> Result<Vec<String>, String> {
    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join(format!("out.{extension}"));
    let mut book = Book::from_file(asset(&format!("input.{extension}"))).unwrap();
    book.replace(
        [],
        [Replace {
            find: Find {
                mode: Mode::Regex,
                match_case,
                entire_cells,
                ..Find::literal(find)
            },
            with: with.to_string(),
        }],
    );
    book.write(&out).map_err(|e| e.to_string())?;
    let workbook = Book::from_file(&out).unwrap().workbook().unwrap();
    let sheet = workbook.sheets.iter().find(|s| s.name == "T").unwrap();
    Ok((0..texts().len() as u32)
        .map(|row| match sheet.cells.get(&(row, 0)).map(|c| &c.value) {
            None | Some(CellValue::Empty) => String::new(),
            Some(CellValue::Text(t)) => t.to_string(),
            // What was typed became a number (`a1` with `^a` replaced by nothing): its text.
            Some(CellValue::Number(n)) => n.to_string(),
            Some(other) => format!("{other:?}"),
        })
        .collect())
}

#[test]
fn every_search_of_the_oracle_leaves_what_libreoffice_leaves() {
    let texts = texts();
    for extension in ["xlsx", "ods"] {
        let mut different = Vec::new();
        let mut checked = 0;
        for case in ORACLE["cases"].as_array().unwrap() {
            if case["format"] != extension {
                continue;
            }
            let (find, with) = (
                case["find"].as_str().unwrap(),
                case["with"].as_str().unwrap(),
            );
            let (match_case, entire) = (
                case["match_case"].as_bool().unwrap(),
                case["entire_cells"].as_bool().unwrap(),
            );
            let want: Vec<String> = case["result"]
                .as_array()
                .unwrap()
                .iter()
                .map(|t| t.as_str().unwrap().to_string())
                .collect();
            checked += 1;
            match replaced(extension, find, with, match_case, entire) {
                Err(why) => different.push(format!("{find:?} -> {with:?}: refused: {why}")),
                Ok(got) => {
                    for ((text, got), want) in texts.iter().zip(&got).zip(&want) {
                        if got != want {
                            different.push(format!(
                                "{find:?} -> {with:?} (case {match_case}, entire {entire}) on {text:?}: got {got:?}, LibreOffice {want:?}"
                            ));
                        }
                    }
                }
            }
        }
        assert!(checked > 500, "{extension}: {checked} searches");
        assert!(
            different.is_empty(),
            "{extension}: {} differences:\n{}",
            different.len(),
            different
                .iter()
                .take(40)
                .cloned()
                .collect::<Vec<_>>()
                .join("\n")
        );
    }
}

/// The result for one text, in both formats (which must agree).
fn one(find: &str, with: &str, text: &str) -> String {
    let at = texts()
        .iter()
        .position(|t| t == text)
        .expect("a text of the fixture");
    let xlsx = replaced("xlsx", find, with, false, false).unwrap()[at].clone();
    assert_eq!(xlsx, replaced("ods", find, with, false, false).unwrap()[at]);
    xlsx
}

#[test]
fn a_caret_alone_replaces_nothing_and_a_dollar_alone_adds_at_the_end() {
    assert_eq!(one("^", "X", "abc"), "abc");
    assert_eq!(one("$", "X", "abc"), "abcX");
    assert_eq!(one("^$", "X", "abc"), "abc");
    assert_eq!(one("$", "[&]", "abc"), "abc[]");
}

#[test]
fn a_pattern_that_can_match_the_empty_text_replaces_what_it_matches_of_it() {
    assert_eq!(one("x*", "X", "abc"), "abc");
    assert_eq!(one("x*", "X", "xax"), "XaX");
    assert_eq!(one("a*", "X", "abab"), "XbXb");
    assert_eq!(one("a*", "X", "ab"), "Xb");
    assert_eq!(one("a*$", "X", "aa"), "X");
    assert_eq!(one("a*$", "X", "ab"), "abX");
}

#[test]
fn the_end_of_the_text_is_not_searched_after_a_hit_or_a_skipped_empty_match() {
    // `(^|$)`: the empty match at the start is skipped, and the search that follows it begins at
    // the end of a one-character text, where none does.
    assert_eq!(one("(^|$)", "X", "a"), "a");
    assert_eq!(one("(^|$)", "X", "ab"), "abX");
    assert_eq!(one("(b|$)", "X", "b"), "X");
    assert_eq!(one("(b|$)", "X", "ba"), "XaX");
}

#[test]
fn a_replacement_is_made_in_the_text_and_the_next_search_goes_on_in_what_is_left() {
    // The second `a` is at the start once the first is gone.
    assert_eq!(one("^a", "", "aaa"), "");
    assert_eq!(one("^a", "", "aab"), "b");
    assert_eq!(one("^a", "X", "aa"), "Xa");
}

#[test]
fn a_character_outside_the_basic_plane_is_two_units_for_the_search_that_goes_on_after_an_empty_match()
 {
    // `LibreOffice` counts in UTF-16 units: after the empty match at the start of the emoji its
    // search begins inside it, so `(^|$)` still finds the end, and `x*` matches the empty text
    // there and finds nothing more.
    let emoji = "\u{1F600}";
    assert_eq!(one("(^|$)", "X", emoji), format!("{emoji}X"));
    assert_eq!(one("x*", "X", emoji), emoji);
    assert_eq!(one("a*", "X", "a\u{1F600}"), "X\u{1F600}");
}

#[test]
fn a_dollar_is_before_the_line_break_that_ends_the_text_and_never_before_another() {
    assert_eq!(one("a$", "X", "a\n"), "X\n");
    assert_eq!(
        one("$", "X", "a\n"),
        "a\n",
        "the empty match before the break is skipped"
    );
    assert_eq!(one("a$", "X", "a\nb"), "a\nb");
    assert_eq!(one("b$", "X", "a\nb"), "a\nX");
    assert_eq!(one("$", "X", "a\nb"), "a\nbX");
}

#[test]
fn a_caret_is_the_start_of_the_text_and_not_of_a_line() {
    assert_eq!(one("^a", "X", "a\nb"), "X\nb");
    assert_eq!(one("^b", "X", "a\nb"), "a\nb");
    assert_eq!(
        one("^.", "X", "\na"),
        "\na",
        "a dot does not match a line break"
    );
}

#[test]
fn the_replacement_of_an_anchored_search_has_the_groups_of_the_match() {
    assert_eq!(one("^(a)(b)", "<$2$1>", "abc"), "<ba>c");
    assert_eq!(one("(^|x)(a)", "<$1|$2>", "xa"), "<x|a>");
    assert_eq!(one("(^|x)(a)", "<$1|$2>", "ab"), "<|a>b");
    assert_eq!(one("(a*)$", "<$1>", "ab"), "ab<>");
    assert_eq!(one("^a", "[&]", "ab"), "[a]b");
}

#[test]
fn a_pattern_that_is_not_reproduced_is_still_refused() {
    for extension in ["xlsx", "ods"] {
        for pattern in ["", "^*", "(a*)*", "\\bfoo", "(?=a)b"] {
            let err = replaced(extension, pattern, "X", false, false).unwrap_err();
            assert!(err.contains("unsupported search"), "{pattern}: {err}");
        }
    }
}

//! What a search finds with a pattern that matches the empty text (`^`, `$`, `x*`, `a*$`,
//! `(^|$)`), and with the empty search text, against `LibreOffice` itself.
//!
//! `tests/assets/find_empty/` has a workbook made with `LibreOffice` 26.8 in both formats
//! (`input.*`, `T!A1:A..`, one text a row, among them texts of several lines) and `oracle.json`, the
//! rows that `Sheet.findAll` reported for each regular expression, `match_case` and `entire_cells`
//! (`make_find_empty_oracle.py`; the patterns are those of the replace oracle). The search runs the
//! same searches and must report the same rows.
//!
//! What it says, in short: `findAll` reports the cells that `replaceAll` would change, so a match of
//! no characters is a hit only at the very end of the text (`$` reports every cell, `^` and `x*` none
//! unless something else matches), and the empty search text finds nothing, in any mode.

use serde_json::Value as Json;
use sheetrs_io::Book;
use sheetrs_protocol::{Find, Mode, Query};
use sheetrs_query::BookQuery;
use std::path::{Path, PathBuf};
use std::sync::LazyLock;

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/find_empty")
        .join(name)
}

static ORACLE: LazyLock<Json> = LazyLock::new(|| {
    serde_json::from_str(&std::fs::read_to_string(asset("oracle.json")).unwrap()).unwrap()
});

/// The 0-based rows of `T!A:A` a search reports.
fn search(extension: &str, find: Find) -> Result<Vec<u64>, String> {
    let query = Query {
        find: Some(find),
        ..Query::default()
    };
    let book = Book::from_file(asset(&format!("input.{extension}"))).unwrap();
    let mut rows: Vec<u64> = book
        .query(&query)
        .map_err(|e| e.to_string())?
        .cells
        .iter()
        .map(|c| c.cell()[1..].parse::<u64>().unwrap() - 1)
        .collect();
    rows.sort_unstable();
    Ok(rows)
}

fn rows(json: &Json) -> Vec<u64> {
    json.as_array()
        .unwrap()
        .iter()
        .map(|r| r.as_u64().unwrap())
        .collect()
}

#[test]
fn every_search_of_the_oracle_reports_the_rows_libreoffice_reports() {
    let texts = ORACLE["texts"].as_array().unwrap();
    for extension in ["xlsx", "ods"] {
        let mut different = Vec::new();
        let mut checked = 0;
        for case in ORACLE["cases"].as_array().unwrap() {
            if case["format"] != extension {
                continue;
            }
            let pattern = case["find"].as_str().unwrap();
            let find = Find {
                mode: Mode::Regex,
                match_case: case["match_case"].as_bool().unwrap(),
                entire_cells: case["entire_cells"].as_bool().unwrap(),
                ..Find::literal(pattern)
            };
            let want = rows(&case["rows"]);
            checked += 1;
            match search(extension, find) {
                Err(why) => different.push(format!("{pattern:?}: refused: {why}")),
                Ok(got) if got != want => different.push(format!(
                    "{pattern:?} (case {}, entire {}): got {:?}, LibreOffice {:?}",
                    case["match_case"],
                    case["entire_cells"],
                    got.iter().map(|&r| &texts[r as usize]).collect::<Vec<_>>(),
                    want.iter().map(|&r| &texts[r as usize]).collect::<Vec<_>>(),
                )),
                Ok(_) => {}
            }
        }
        assert!(checked > 400, "{extension}: {checked} searches");
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

#[test]
fn the_empty_search_text_finds_nothing_in_every_mode() {
    for extension in ["xlsx", "ods"] {
        for case in ORACLE["empty"].as_array().unwrap() {
            if case["format"] != extension {
                continue;
            }
            // `LibreOffice` finds nothing.
            assert!(rows(&case["rows"]).is_empty(), "{case}");
            let mode = match case["mode"].as_str().unwrap() {
                "regex" => Mode::Regex,
                "wildcards" => Mode::Wildcards,
                _ => Mode::Literal,
            };
            let find = Find {
                mode,
                match_case: case["match_case"].as_bool().unwrap(),
                entire_cells: case["entire_cells"].as_bool().unwrap(),
                ..Find::literal("")
            };
            assert_eq!(search(extension, find), Ok(vec![]), "{case}");
        }
    }
}

#[test]
fn the_empty_search_text_is_answered_before_anything_is_read() {
    // No error either for a sheet that is not there, or for a file with nothing to read.
    let query = Query {
        scope: vec![sheetrs_protocol::Scope::Sheet("Nowhere".into())],
        find: Some(Find::literal("")),
        ..Query::default()
    };
    let book = Book::from_file(asset("input.xlsx")).unwrap();
    assert!(book.query(&query).unwrap().cells.is_empty());
}

/// The texts of the fixture that a search reports.
fn found(pattern: &str, entire: bool) -> Vec<String> {
    let texts = ORACLE["texts"].as_array().unwrap();
    let rows = search(
        "xlsx",
        Find {
            mode: Mode::Regex,
            entire_cells: entire,
            ..Find::literal(pattern)
        },
    )
    .unwrap();
    rows.iter()
        .map(|&r| texts[r as usize].as_str().unwrap().to_string())
        .collect()
}

#[test]
fn an_empty_match_is_a_hit_only_at_the_very_end_of_the_text() {
    // The end of the text: every text that does not end in a line break (the empty match before
    // that one is skipped, and no search begins at the end).
    let end = found("$", false);
    for hit in ["a", "abc", "a\nb", "\u{1F600}", "a\u{1F600}b"] {
        assert!(end.iter().any(|t| t == hit), "{hit:?}");
    }
    for miss in ["a\n", "\n", "a\u{2028}", "a\n\n"] {
        assert!(!end.iter().any(|t| t == miss), "{miss:?}");
    }
    // The start of the text, and the empty match of `x*` in the middle, are not: only what the
    // pattern matches with characters is.
    assert!(found("^", false).is_empty());
    assert_eq!(found("x*", false), found("x", false));
    assert_eq!(found("^a*", false), found("^a", false));
    // With `entire_cells` the first match must cover the text: an empty one does not.
    assert!(found("$", true).is_empty());
    assert!(found("^$", true).is_empty());
    assert!(found("(^|$)", true).is_empty());
}

#[test]
fn a_pattern_that_is_not_in_the_subset_is_still_refused() {
    for pattern in ["^*", "(^)+a", "^+", "(a*)*", "\\b", "(?=a)", "a**"] {
        let refused = search(
            "xlsx",
            Find {
                mode: Mode::Regex,
                ..Find::literal(pattern)
            },
        );
        assert!(
            refused
                .as_ref()
                .is_err_and(|e| e.contains("unsupported search")),
            "{pattern}: {refused:?}"
        );
    }
}

//! The copies of a shared formula that would read outside the sheet, against `LibreOffice` 26.8.
//!
//! `tests/assets/write_to/shared_beyond.xlsx` has, in the last columns of the sheet (`XFB:XFD` of `Calc`), nine groups of
//! three cells that share a formula (`XFC2+1`, `SUM(XFB1:XFC1)`, `Data!XFC1`, ...): the copy in `XFD` reads column `XFE`,
//! which is not a cell. `shared_beyond.json` (`make_shared_oracle.py`) is what `LibreOffice` shows when the file is
//! SAVED AND OPENED AGAIN (project rule, `CONTRIBUTING.md`; what it shows when it opens the fixture is its memory, `#REF!`
//! with a text such as `=#REF!2+1`, recorded as `first_load` and never compared): the value of each cell (`reread`) and
//! the text it saves in the XML (`saved`, a bare `#REF!`). It is an XLSX case only: an ODS has no shared formulas.
//!
//! The workbook that `SheetRS` writes must show the same values and hold the same texts, in both modes: with
//! `calculate_all` (every formula calculated) and by default (the cache: the fixture caches the results `LibreOffice` calculates,
//! errors included, and `LibreOffice` calculates on load the formula of a name it cannot resolve whose cached result is an
//! error, which is what the default mode of `SheetRS` does too).

use serde_json::{Value as Json, json};
use sheetrs_edit::{BookEdit, Calculation};
use sheetrs_io::Book;
use sheetrs_query::{BookQuery, RangeOptions};
use std::path::Path;

const ROWS: [u32; 9] = [1, 3, 4, 5, 6, 7, 8, 9, 10];
const COLUMNS: [(&str, u32); 3] = [("XFB", 16_381), ("XFC", 16_382), ("XFD", 16_383)];

fn assets() -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/assets/write_to")
}

fn oracle() -> Json {
    serde_json::from_str(&std::fs::read_to_string(assets().join("shared_beyond.json")).unwrap())
        .unwrap()
}

/// What `LibreOffice` shows, as text, against what `SheetRS` has in the cell: a number, a boolean or an error.
fn same_value(shown: &str, mine: &Json) -> bool {
    match mine {
        Json::Number(n) => shown
            .parse::<f64>()
            .is_ok_and(|s| (s - n.as_f64().unwrap()).abs() < 1e-9),
        Json::Bool(b) => shown == if *b { "TRUE" } else { "FALSE" },
        Json::Object(o) => o.get("error").and_then(Json::as_str) == Some(shown),
        Json::String(t) => t == shown,
        _ => false,
    }
}

/// The workbook after a calculation with an input (the whole workbook is the answer), in a mode.
fn written(calculate_all: bool) -> Vec<u8> {
    let mut book = Book::from_file(assets().join("shared_beyond.xlsx"))
        .unwrap()
        .with_calculate_all(calculate_all);
    BookEdit::calculate(
        &mut book,
        Calculation {
            inputs: json!({"Data!A1": 5}).as_object().cloned().unwrap(),
            ..Calculation::default()
        },
    );
    let mut bytes = Vec::new();
    book.write_to(&mut bytes).unwrap();
    bytes
}

fn check(bytes: &[u8], what: &str) {
    let oracle = oracle();
    let book = Book::from_bytes(bytes.to_vec()).unwrap();
    let blocks = book
        .range(["Calc!XFB1:XFD10"], &RangeOptions::default())
        .unwrap();
    let values = blocks[0].values.as_ref().unwrap();
    let workbook = book.workbook().unwrap();
    let calc = workbook.sheets.iter().find(|s| s.name == "Calc").unwrap();
    let mut wrong = Vec::new();
    for row in ROWS {
        for (letters, col) in COLUMNS {
            let cell = format!("{letters}{row}");
            let expected = &oracle["cells"][&cell];
            let mine = values[(row - 1) as usize][(col - 16_381) as usize]
                .clone()
                .unwrap_or(Json::Null);
            let shown = expected["reread"].as_str().unwrap();
            if !same_value(shown, &mine) {
                wrong.push(format!("{cell}: value {mine}, LibreOffice {shown}"));
            }
            let text = calc
                .cells
                .iter()
                .find(|(at, _)| (at.0, at.1) == (row - 1, col))
                .and_then(|(_, c)| c.formula.clone());
            let saved = expected["saved"].as_str().unwrap();
            if text.as_deref().map(|t| t.trim_start_matches('=')) != Some(saved) {
                wrong.push(format!("{cell}: formula {text:?}, LibreOffice {saved:?}"));
            }
        }
    }
    assert!(wrong.is_empty(), "{what}: {wrong:#?}");
}

#[test]
fn calculate_all_gives_what_libreoffice_shows_of_the_file_it_saved() {
    check(&written(true), "calculate_all");
}

#[test]
fn the_default_mode_gives_what_libreoffice_shows_of_the_file_it_saved() {
    check(&written(false), "the cache");
}

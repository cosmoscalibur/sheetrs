//! A replacement that types a date writes it in the date system of the workbook: the three
//! `LibreOffice` Calc has (1899-12-30, 1900-01-01 and 1904-01-01). The files are
//! `sheetrs-io/tests/assets/dates*.{xlsx,ods}`, whose first column holds the same serials in the three
//! systems (45366 is 2024-03-15 in the usual one, and another day in the others) with the format
//! `yyyy-mm-dd`: typing `2024-03-20` over a cell is the serial 45371 minus the days the null date is
//! after 1899-12-30 (0, 2 or 1462), which is where each system numbers that day.

use sheetrs_edit::BookEdit;
use sheetrs_io::Book;
use sheetrs_io::reader::CellValue;
use sheetrs_protocol::{Find, Mode, Replace, Scope, SearchIn};
use sheetrs_query::BookQuery;
use std::path::PathBuf;

fn asset(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-io/tests/assets")
        .join(name)
}

/// The cell that holds the serial 45366 shows a date; typing `2024-03-20` over it leaves the serial of
/// that day in the system of the file, `shift` days after the one of the usual system.
fn typed_over(file: &str, shift: f64) {
    let extension = file.rsplit('.').next().unwrap();
    let source = Book::from_file(asset(file)).unwrap();
    let before = Book::from_file(asset(file)).unwrap().workbook().unwrap();
    let (row, col) = before
        .sheet("Dates")
        .unwrap()
        .cells
        .iter()
        .find_map(|((row, col), cell)| {
            (col == 0 && matches!(cell.value, CellValue::Number(n) if n == 45366.0))
                .then_some((row, col))
        })
        .unwrap_or_else(|| panic!("{file}: no cell holds the serial 45366"));
    // What the cell shows, in the date system of the file (checked against LibreOffice by the tests
    // of `sheetrs-query`).
    let query: sheetrs_protocol::Query = serde_json::from_value(serde_json::json!({
        "properties": ["display"],
        "scope": [{"sheet": "Dates"}]
    }))
    .unwrap();
    let shown = source
        .query(&query)
        .unwrap()
        .cells
        .into_iter()
        .find(|c| c.row == row + 1 && c.col == col + 1)
        .and_then(|c| c.display)
        .unwrap();
    assert!(
        shown.starts_with("20") && shown.len() == 10,
        "{file}: {shown}"
    );
    let find = Find {
        text: shown,
        mode: Mode::Literal,
        match_case: false,
        entire_cells: true,
        search_in: SearchIn::Value,
        include_formulas: false,
        formatted: false,
        filtered: false,
    };
    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join(format!("typed.{extension}"));
    let mut editor = source.clone();
    editor.replace(
        vec![Scope::Sheet("Dates".into())],
        [Replace {
            find,
            with: "2024-03-20".into(),
        }],
    );
    let report = editor.write(&out).unwrap();
    assert!(report.skipped.is_empty(), "{file}: {:?}", report.skipped);
    let after = Book::from_file(&out).unwrap().workbook().unwrap();
    let cell = after.sheet("Dates").unwrap().cell(row, col).unwrap();
    assert!(
        matches!(cell.value, CellValue::Number(n) if n == 45371.0 - shift),
        "{file}: {:?}",
        cell.value
    );
}

#[test]
fn a_date_typed_in_the_default_system_is_numbered_from_1899_12_30() {
    typed_over("dates.xlsx", 0.0);
    typed_over("dates.ods", 0.0);
}

#[test]
fn a_date_typed_in_the_starcalc_system_is_numbered_from_1900_01_01() {
    typed_over("dates_1900.ods", 2.0);
}

#[test]
fn a_date_typed_in_the_1904_system_is_numbered_from_1904_01_01() {
    typed_over("dates_1904.xlsx", 1462.0);
    typed_over("dates_1904.ods", 1462.0);
}

/// Another null date is an error that names it, before anything is typed into the workbook.
#[test]
fn a_workbook_with_another_null_date_is_not_edited() {
    let source = Book::from_file(asset("other_null_date.ods")).unwrap();
    let out = tempfile::tempdir().unwrap().path().join("never.ods");
    let mut editor = source.clone();
    editor.replace(
        Vec::new(),
        [Replace {
            find: Find {
                text: "x".into(),
                mode: Mode::Literal,
                match_case: false,
                entire_cells: false,
                search_in: SearchIn::Value,
                include_formulas: false,
                formatted: false,
                filtered: false,
            },
            with: "y".into(),
        }],
    );
    let error = editor.write(&out).unwrap_err().to_string();
    assert!(error.contains("1950-06-15"), "{error}");
}

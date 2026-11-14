//! What a conversion of formulas to values converts to: what the formulas calculate, not what the file caches
//! (the result a file caches is not a goal of the project, `docs/approved-differences.md`), unless the run uses
//! the cache. The workbooks of `tests/assets/use_cache/` cache results the engine would not give (`from-excel`
//! where the engine gives `t`, a `ROW()` that caches 99), so a value that came from the cache is told from one
//! that was calculated, in each format.

use sheetrs_edit::{BookEdit, CONVERTED_FROM_CACHE};
use sheetrs_io::Book;
use sheetrs_io::reader::{CellValue, Workbook};
use std::path::{Path, PathBuf};

const EXTENSIONS: [&str; 2] = ["xlsx", "ods"];

fn asset(folder: &str, name: &str, extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(folder)
        .join(format!("{name}.{extension}"))
}

fn at(workbook: &Workbook, sheet: &str, row: u32, col: u32) -> (Option<String>, CellValue) {
    let cell = &workbook
        .sheets
        .iter()
        .find(|s| s.name == sheet)
        .unwrap()
        .cells[&(row, col)];
    (
        cell.formula.as_deref().map(str::to_string),
        cell.value.clone(),
    )
}

fn number(n: f64) -> CellValue {
    CellValue::Number(n)
}

/// `book` of `use_cache/`, edited by `edit`, as the workbook that `write_to` gives.
fn edited(
    extension: &str,
    use_cache: bool,
    edit: impl Fn(&mut Book),
) -> (Workbook, Vec<sheetrs_protocol::Diagnostic>) {
    let mut book = Book::from_file(asset("use_cache", "book", extension))
        .unwrap()
        .with_use_cache(use_cache);
    edit(&mut book);
    let mut bytes = Vec::new();
    let report = book.write_to(&mut bytes).unwrap();
    (
        Book::from_bytes(bytes).unwrap().workbook().unwrap(),
        report.diagnostics,
    )
}

#[test]
fn a_conversion_gives_what_the_formulas_calculate_and_not_what_the_file_caches() {
    for extension in EXTENSIONS {
        let (workbook, diagnostics) = edited(extension, false, |book| {
            book.formula_to_value(["Calc!A:A".parse().unwrap()]);
        });
        assert!(diagnostics.is_empty(), "{extension}: {diagnostics:?}");
        // `rate*2` caches 140, `IF(Data!B1<>"","t",0)` caches `from-excel`, `ROW()` caches 99.
        assert_eq!(
            at(&workbook, "Calc", 5, 0),
            (None, number(14.0)),
            "{extension}"
        );
        assert_eq!(
            at(&workbook, "Calc", 6, 0),
            (None, CellValue::Text("t".into())),
            "{extension}"
        );
        assert_eq!(
            at(&workbook, "Calc", 7, 0),
            (None, number(8.0)),
            "{extension}"
        );
        assert_eq!(
            at(&workbook, "Calc", 0, 0),
            (None, number(210.0)),
            "{extension}"
        );
    }
}

#[test]
fn a_conversion_calculates_the_cells_it_names_and_leaves_the_others_as_the_file_has_them() {
    for extension in EXTENSIONS {
        let (workbook, _) = edited(extension, false, |book| {
            book.formula_to_value(["Calc!A:A".parse().unwrap()]);
        });
        // `B2` is not converted and nothing reached it: its formula and its cached 777 stay.
        let (formula, value) = at(&workbook, "Calc", 1, 1);
        assert!(formula.is_some(), "{extension}");
        assert_eq!(value, number(777.0), "{extension}");
    }
}

#[test]
fn with_the_cache_a_conversion_keeps_the_results_the_file_caches() {
    for extension in EXTENSIONS {
        let (workbook, _) = edited(extension, true, |book| {
            book.formula_to_value(["Calc!A:A".parse().unwrap()]);
        });
        assert_eq!(
            at(&workbook, "Calc", 5, 0),
            (None, number(140.0)),
            "{extension}"
        );
        assert_eq!(
            at(&workbook, "Calc", 6, 0),
            (None, CellValue::Text("from-excel".into())),
            "{extension}"
        );
        assert_eq!(
            at(&workbook, "Calc", 7, 0),
            (None, number(99.0)),
            "{extension}"
        );
    }
}

#[test]
fn a_conversion_with_a_deletion_converts_the_workbook_as_it_is_before_the_cells_move() {
    for extension in EXTENSIONS {
        let (workbook, _) = edited(extension, false, |book| {
            book.formula_to_value(["Calc!A:A".parse().unwrap()]);
            book.delete_rows(["Data!2".parse().unwrap()]);
        });
        // `SUM(Data!A1:A6)` is 210 and `Data!A2*2` is 40 where the row is still there; the deletion
        // that follows changes what the formulas that are left give, not these.
        assert_eq!(
            at(&workbook, "Calc", 0, 0),
            (None, number(210.0)),
            "{extension}"
        );
        assert_eq!(
            at(&workbook, "Calc", 1, 0),
            (None, number(40.0)),
            "{extension}"
        );
        assert_eq!(
            at(&workbook, "Calc", 5, 0),
            (None, number(14.0)),
            "{extension}"
        );
    }
}

#[test]
fn a_conversion_of_a_workbook_the_engine_cannot_calculate_takes_the_cache_and_says_so() {
    // The book has a formula that reads another workbook, which the engine does not read.
    for extension in EXTENSIONS {
        let mut book = Book::from_file(asset("book", "book", extension)).unwrap();
        book.formula_to_value(["Data".parse().unwrap()]);
        let mut bytes = Vec::new();
        let report = book.write_to(&mut bytes).unwrap();
        let said: Vec<_> = report
            .diagnostics
            .iter()
            .filter(|d| d.code == CONVERTED_FROM_CACHE)
            .collect();
        assert_eq!(said.len(), 1, "{extension}: {:?}", report.diagnostics);
        // The cells are values, the results the file caches.
        let workbook = Book::from_bytes(bytes).unwrap().workbook().unwrap();
        assert!(
            workbook
                .sheets
                .iter()
                .find(|s| s.name == "Data")
                .unwrap()
                .cells
                .values()
                .all(|cell| cell.formula.is_none()),
            "{extension}"
        );
    }
}

#[test]
fn a_conversion_of_a_workbook_the_engine_cannot_calculate_and_whose_results_are_stale_stops() {
    // A deletion before the conversion leaves the cached results stale (a replacement does too): there
    // are no results to take, so the run stops, as it does for a conversion that needs the engine.
    for extension in EXTENSIONS {
        let mut book = Book::from_file(asset("book", "book", extension)).unwrap();
        let replacement = sheetrs_protocol::Replace {
            find: sheetrs_protocol::Find::literal("Total"),
            with: "Sum".to_string(),
        };
        book.replace(["Summary".into()], [replacement]);
        book.formula_to_value(["Data".parse().unwrap()]);
        let error = book.write_to(&mut Vec::new()).unwrap_err();
        assert!(
            error.to_string().contains("could not be calculated"),
            "{extension}: {error}"
        );
    }
}

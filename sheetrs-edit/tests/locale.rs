//! The locale of the `Book` reaches the calculation of the plan: a `calculate` step, and the
//! recalculation that follows an edit that leaves results stale. The oracle is
//! `semantics_es-CO.{xlsx,ods}` of the compiler: made by a `LibreOffice` running with the `es-CO`
//! regional setting, with the cases of issues #84 and #106 among its rows.

use sheetrs_edit::{BookEdit, Calculation};
use sheetrs_io::Book;
use sheetrs_io::reader::CellValue;
use std::path::{Path, PathBuf};

const EXTENSIONS: [&str; 2] = ["xlsx", "ods"];

/// The cases of issues #84 and #106: the formula (the label of its row in the oracle), as `es-CO` and
/// as `en-US` show it.
const CASES: [(&str, &str, &str); 3] = [
    (
        r#"TEXT(350*49799,"[$$-en-US]#.###.##0")"#,
        "$17.429.650",
        "$17429650.000.000",
    ),
    (
        r#""Componente inflacionario "&(0.5543*100)&"%""#,
        "Componente inflacionario 55,43%",
        "Componente inflacionario 55.43%",
    ),
    (
        r#"TEXT(0,"[$$-en-US]#,##0;([$$-en-US]#,##0)")"#,
        "$,000",
        "$0",
    ),
];

fn oracle(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-compiler/tests/assets")
        .join(format!("semantics_es-CO.{extension}"))
}

fn book(extension: &str, locale: Option<&str>) -> Book {
    let book = Book::from_file(oracle(extension)).unwrap();
    match locale {
        Some(tag) => book.with_locale(tag.parse().unwrap()),
        None => book,
    }
}

/// The 0-based row of the `Calc` sheet of the oracle whose label (column B) is `formula`.
fn row_of(extension: &str, formula: &str) -> u32 {
    let workbook = Book::from_file(oracle(extension))
        .unwrap()
        .workbook()
        .unwrap();
    let calc = workbook.sheets.iter().find(|s| s.name == "Calc").unwrap();
    calc.cells
        .iter()
        .find_map(|(at, cell)| {
            (at.1 == 1 && matches!(&cell.value, CellValue::Text(t) if &**t == formula))
                .then_some(at.0)
        })
        .unwrap_or_else(|| panic!("no row for {formula}"))
}

fn text_at(path: &Path, row: u32) -> String {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let calc = workbook.sheets.iter().find(|s| s.name == "Calc").unwrap();
    match calc
        .cells
        .iter()
        .find(|(at, _)| (at.0, at.1) == (row, 0))
        .map(|(_, cell)| cell.value.clone())
    {
        Some(CellValue::Text(t)) => t.to_string(),
        other => panic!("A{} is {other:?}, not a text", row + 1),
    }
}

#[test]
fn a_calculate_step_uses_the_locale_of_the_book() {
    // The XLSX takes the one given to the book, and the ODS says it in the file; a book given
    // another one has that one in both.
    for extension in EXTENSIONS {
        for (given, spanish) in [(Some("es-CO"), true), (Some("en-US"), false)] {
            for (formula, es, en) in CASES {
                let asked = format!("Calc!A{}", row_of(extension, formula) + 1);
                let mut book = book(extension, given);
                book.calculate(Calculation {
                    ranges: Some(vec![asked.clone()]),
                    answer: true,
                    ..Calculation::default()
                });
                let report = book.write_to(std::io::sink()).unwrap();
                let answer = report.value_ranges.unwrap().to_json();
                assert_eq!(
                    answer["valueRanges"][0]["values"][0][0],
                    if spanish { es } else { en },
                    "{extension} {given:?}: {formula}"
                );
            }
        }
    }
    // An ODS with no locale given has the one of the file: es-CO.
    let (formula, es, _) = CASES[0];
    let asked = format!("Calc!A{}", row_of("ods", formula) + 1);
    let mut ods = book("ods", None);
    ods.calculate(Calculation {
        ranges: Some(vec![asked]),
        answer: true,
        ..Calculation::default()
    });
    let answer = ods
        .write_to(std::io::sink())
        .unwrap()
        .value_ranges
        .unwrap()
        .to_json();
    assert_eq!(answer["valueRanges"][0]["values"][0][0], es);
}

#[test]
fn the_recalculation_that_follows_an_edit_uses_the_locale_of_the_book() {
    let folder = tempfile::tempdir().unwrap();
    for extension in EXTENSIONS {
        for (given, spanish) in [(Some("es-CO"), true), (Some("en-US"), false)] {
            // Deleting a sheet that the formulas of the cases do not read leaves the results stale.
            let mut book = book(extension, given);
            book.delete_sheets(["Wide"]);
            let out = folder.path().join(format!("{spanish}.{extension}"));
            book.write(&out).unwrap();
            for (formula, es, en) in CASES {
                let row = row_of(extension, formula);
                assert_eq!(
                    text_at(&out, row),
                    if spanish { es } else { en },
                    "{extension} {given:?}: {formula}"
                );
            }
        }
    }
}

//! Rows hidden and rows deleted in one plan: every row of the plan is a row of the input, and the formulas
//! that read the hidden rows are calculated again where the rows are now, in an XLSX and in an ODS.

mod common;

use common::{Cell, EXTENSIONS, formula, number, sheet, workbook};
use sheetrs_edit::BookEdit;
use sheetrs_io::Book;

/// `Data!A4:A5` holds 4 and 5, and `Calc!A1` caches the `SUBTOTAL` that adds them (9).
fn book(extension: &str) -> std::path::PathBuf {
    let dir = tempfile::tempdir().unwrap().keep();
    workbook(
        &dir,
        extension,
        &[
            sheet(
                "Data",
                vec![("A4", Cell::Number(4.0)), ("A5", Cell::Number(5.0))],
            ),
            sheet(
                "Calc",
                vec![(
                    "A1",
                    Cell::Formula(
                        "SUBTOTAL(109,Data!A4:A5)",
                        "SUBTOTAL(109;[$Data.A4:.A5])",
                        9.0,
                    ),
                )],
            ),
        ],
        &[],
    )
}

#[test]
fn a_subtotal_skips_the_row_hidden_in_the_same_plan_that_deletes_an_earlier_row() {
    for extension in EXTENSIONS {
        let input = book(extension);
        let out = input.with_file_name(format!("out.{extension}"));
        Book::from_file(&input)
            .unwrap()
            .hide_rows(["Data!5".parse().unwrap()])
            .delete_rows(["Data!1".parse().unwrap()])
            .write(&out)
            .unwrap();
        assert_eq!(
            formula(&out, "Calc", "A1").as_deref(),
            Some("SUBTOTAL(109,Data!A3:A4)"),
            "{extension}"
        );
        assert!(
            (number(&out, "Calc", "A1") - 4.0).abs() < 1e-9,
            "{extension}: {}",
            number(&out, "Calc", "A1")
        );
    }
}

#[test]
fn a_subtotal_skips_the_row_hidden_in_a_plan_with_no_deletion() {
    for extension in EXTENSIONS {
        let input = book(extension);
        let out = input.with_file_name(format!("out.{extension}"));
        Book::from_file(&input)
            .unwrap()
            .hide_rows(["Data!5".parse().unwrap()])
            .write(&out)
            .unwrap();
        assert!(
            (number(&out, "Calc", "A1") - 4.0).abs() < 1e-9,
            "{extension}: {}",
            number(&out, "Calc", "A1")
        );
    }
}

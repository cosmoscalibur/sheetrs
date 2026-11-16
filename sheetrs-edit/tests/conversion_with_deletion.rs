//! A conversion of formulas to values and a deletion of rows in one pass over the file: the formulas the pass
//! does not convert keep the text a deletion gives them (their references move with the rows) when the
//! engine writes the result it calculated for them, in an XLSX and in an ODS.

mod common;

use common::{Cell, EXTENSIONS, formula, number, sheet, workbook};
use sheetrs_edit::BookEdit;
use sheetrs_io::{Book, Volatile};
use std::path::PathBuf;

/// The clock, in every test.
const NOW: f64 = 100.0;

/// `Data!C` is converted and `Data!1` is deleted, with the clock at `NOW`.
fn converted_and_deleted(extension: &str, cells: Vec<(&'static str, Cell)>) -> PathBuf {
    let dir = tempfile::tempdir().unwrap().keep();
    let input = workbook(&dir, extension, &[sheet("Data", cells)], &[]);
    let out = dir.join(format!("out.{extension}"));
    Book::from_file(&input)
        .unwrap()
        .with_volatile(Volatile {
            now: Some(NOW),
            seed: None,
        })
        .formula_to_value(["Data!C".parse().unwrap()])
        .delete_rows(["Data!1".parse().unwrap()])
        .write(&out)
        .unwrap();
    out
}

/// `Data!A2` and `A3` hold 2 and 3, `B2` reads the clock and `A2` (it caches 5), and `C2` is `1+1`.
fn clock_and_a_cell() -> Vec<(&'static str, Cell)> {
    vec![
        ("A2", Cell::Number(2.0)),
        ("A3", Cell::Number(3.0)),
        ("B2", Cell::Formula("NOW()+A2", "NOW()+[.A2]", 5.0)),
        ("C2", Cell::Formula("1+1", "1+1", 2.0)),
    ]
}

#[test]
fn a_formula_outside_the_conversion_moves_with_the_deleted_row() {
    for extension in EXTENSIONS {
        let out = converted_and_deleted(extension, clock_and_a_cell());
        assert_eq!(
            formula(&out, "Data", "B1").as_deref(),
            Some("NOW()+A1"),
            "{extension}"
        );
    }
}

#[test]
fn a_formula_outside_the_conversion_gives_the_result_of_the_cell_it_reads_now() {
    for extension in EXTENSIONS {
        let out = converted_and_deleted(extension, clock_and_a_cell());
        assert!(
            (number(&out, "Data", "B1") - (NOW + 2.0)).abs() < 1e-9,
            "{extension}: {}",
            number(&out, "Data", "B1")
        );
    }
}

#[test]
fn the_converted_formula_becomes_its_value_in_the_row_it_moved_to() {
    for extension in EXTENSIONS {
        let out = converted_and_deleted(extension, clock_and_a_cell());
        assert_eq!(formula(&out, "Data", "C1"), None, "{extension}");
        assert!(
            (number(&out, "Data", "C1") - 2.0).abs() < 1e-9,
            "{extension}"
        );
    }
}

/// A shared group `B2:B4` of `NOW()+A2` (an XLSX; an ODS has formulas of their own) reading `A2:A4`.
fn shared_group(extension: &str) -> Vec<(&'static str, Cell)> {
    let ods = |text: &'static str| Cell::Formula("", text, 0.0);
    let (b2, b3, b4) = if extension == "xlsx" {
        (
            Cell::Raw(
                r#"<f t="shared" ref="B2:B4" si="0">NOW()+A2</f><v>0</v>"#,
                "",
            ),
            Cell::Raw(r#"<f t="shared" si="0"/><v>0</v>"#, ""),
            Cell::Raw(r#"<f t="shared" si="0"/><v>0</v>"#, ""),
        )
    } else {
        (
            ods("NOW()+[.A2]"),
            ods("NOW()+[.A3]"),
            ods("NOW()+[.A4]"),
        )
    };
    vec![
        ("A2", Cell::Number(2.0)),
        ("A3", Cell::Number(3.0)),
        ("A4", Cell::Number(4.0)),
        ("B2", b2),
        ("B3", b3),
        ("B4", b4),
        ("C2", Cell::Formula("1+1", "1+1", 2.0)),
    ]
}

#[test]
fn a_shared_group_outside_the_conversion_moves_with_the_deleted_row() {
    for extension in EXTENSIONS {
        let out = converted_and_deleted(extension, shared_group(extension));
        for (cell, text, wanted) in [
            ("B1", "NOW()+A1", NOW + 2.0),
            ("B2", "NOW()+A2", NOW + 3.0),
            ("B3", "NOW()+A3", NOW + 4.0),
        ] {
            assert_eq!(
                formula(&out, "Data", cell).as_deref(),
                Some(text),
                "{extension} {cell}"
            );
            assert!(
                (number(&out, "Data", cell) - wanted).abs() < 1e-9,
                "{extension} {cell}: {}",
                number(&out, "Data", cell)
            );
        }
    }
}

/// An array `E2:E3` of `A2:A3+NOW()`.
fn array(extension: &str) -> Vec<(&'static str, Cell)> {
    let (e2, e3) = if extension == "xlsx" {
        (
            Cell::Raw(r#"<f t="array" ref="E2:E3">A2:A3+NOW()</f><v>0</v>"#, ""),
            Cell::Raw(r"<v>0</v>", ""),
        )
    } else {
        (
            Cell::Raw(
                "",
                r#"<table:table-cell table:formula="of:=[.A2:.A3]+NOW()" table:number-matrix-columns-spanned="1" table:number-matrix-rows-spanned="2" office:value-type="float" office:value="0"><text:p>0</text:p></table:table-cell>"#,
            ),
            Cell::Raw(
                "",
                r#"<table:covered-table-cell office:value-type="float" office:value="0"><text:p>0</text:p></table:covered-table-cell>"#,
            ),
        )
    };
    vec![
        ("A2", Cell::Number(2.0)),
        ("A3", Cell::Number(3.0)),
        ("E2", e2),
        ("E3", e3),
        ("C2", Cell::Formula("1+1", "1+1", 2.0)),
    ]
}

#[test]
fn an_array_outside_the_conversion_moves_with_the_deleted_row() {
    for extension in EXTENSIONS {
        let out = converted_and_deleted(extension, array(extension));
        assert_eq!(
            formula(&out, "Data", "E1").as_deref(),
            Some("A1:A2+NOW()"),
            "{extension}"
        );
        assert!(
            (number(&out, "Data", "E1") - (NOW + 2.0)).abs() < 1e-9,
            "{extension}: {}",
            number(&out, "Data", "E1")
        );
        assert!(
            (number(&out, "Data", "E2") - (NOW + 3.0)).abs() < 1e-9,
            "{extension}: {}",
            number(&out, "Data", "E2")
        );
    }
}

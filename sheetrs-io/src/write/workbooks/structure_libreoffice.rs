//! A command on the XLSX and on the ODS that `LibreOffice` wrote from one workbook
//! (`sheetrs-edit/scripts/gen_structure_fixtures.py`), each file judged on its own: the formulas it
//! ends with are what `LibreOffice` 26.8 gave for the same deletion in that file, asked through UNO
//! (`sheetrs-edit/scripts/compare_structure_with_libreoffice.py` makes the comparison live; the
//! expectations here are what it found).

use crate::Book;
use crate::write::{ValueScope, WorkbookModifications, modify_workbook};
use std::collections::HashSet;
use std::path::{Path, PathBuf};

fn asset(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(format!(
        "../sheetrs-edit/tests/assets/structure/book.{extension}"
    ))
}

fn scopes(specs: &[&str]) -> Vec<ValueScope> {
    specs.iter().map(|spec| spec.parse().unwrap()).collect()
}

/// The formulas of column A of sheet `F` after `mods`, as the reader gives them. A reference that
/// went reads `T!#REF!` in an XLSX and `#REF!` in an ODS: both are `#REF!`.
fn run(extension: &str, mods: &WorkbookModifications) -> Vec<String> {
    let dir = tempfile::tempdir().unwrap();
    let output = dir.path().join(format!("out.{extension}"));
    modify_workbook(&asset(extension), &output, mods).unwrap();
    let book = Book::from_file(&output).unwrap().workbook().unwrap();
    let sheet = book.sheets.iter().find(|s| s.name == "F").unwrap();
    let mut cells: Vec<_> = sheet.cells.iter().filter(|(at, _)| at.1 == 0).collect();
    cells.sort_by_key(|(at, _)| at.0);
    cells
        .into_iter()
        .filter_map(|(_, cell)| cell.formula.as_deref())
        .map(|f| f.replace("T!#REF!", "#REF!"))
        .collect()
}

/// The formula with each range of one cell (`A1:A1`) as that cell: `LibreOffice` writes the 3D
/// reference of one cell as the cell in an XLSX and as a range in an ODS.
fn one_cell_ranges(formula: &str) -> String {
    let range = regex::Regex::new(r"(\$?[A-Z]+\$?[0-9]+):(\$?[A-Z]+\$?[0-9]+)").unwrap();
    range
        .replace_all(formula, |c: &regex::Captures| {
            if c[1] == c[2] {
                c[1].to_string()
            } else {
                c[0].to_string()
            }
        })
        .into_owned()
}

/// The expectation as a file spells it: a range of one cell is that cell (`LibreOffice` writes the
/// reference of one cell in a 3D reference as a cell in an XLSX and as a range in an ODS).
fn norm(expected: &[&str]) -> Vec<String> {
    expected.iter().map(|f| one_cell_ranges(f)).collect()
}

/// Runs `mods` on each format and hands the formulas it ends with, as that file gives them, to
/// `check`, with the format for the message.
fn each(mods: &WorkbookModifications, check: impl Fn(&str, &[String])) {
    for extension in ["xlsx", "ods"] {
        let formulas: Vec<String> = run(extension, mods)
            .iter()
            .map(|f| one_cell_ranges(f))
            .collect();
        check(extension, &formulas);
    }
}

#[test]
fn rows_deleted_give_what_libreoffice_gives() {
    let mods = WorkbookModifications {
        delete_rows: scopes(&["T!4:5"]),
        ..Default::default()
    };
    // `Rows.removeByIndex(3, 2)` in LibreOffice; a reference that goes is `#REF!`.
    each(&mods, |extension, f| {
        assert_eq!(
            f,
            norm(&[
                "T!A2",
                "#REF!",
                "#REF!",
                "T!A4",
                "SUM(T!A2:A3)",
                "SUM(T!A2:A3)",
                "SUM(#REF!)",
                "SUM(T!A4:A6)",
                "SUM(T!A2:A6)",
                "SUM(T!A4:A4)",
                "SUM(T!A:A)",
                "SUM(#REF!)",
                "SUM(T!3:4)",
                "SUM(T!1:3)",
                "T!$A$4",
                "SUM(T!A4:B5)",
                "NameOne+1",
                "NameTwo",
                "SUM(T!B2:D2)",
                "SUM(T!C1:C8)",
                "SUM(T!B:C)",
            ]),
            "{extension}"
        );
    });
}

#[test]
fn several_scopes_read_the_rows_of_the_input() {
    let mods = WorkbookModifications {
        // Rows 2 and 7-8 of the input, in any order.
        delete_rows: scopes(&["T!7:8", "T!2"]),
        ..Default::default()
    };
    each(&mods, |extension, f| {
        assert_eq!(
            f,
            norm(&[
                "#REF!",
                "T!A3",
                "T!A4",
                "T!A5",
                "SUM(T!A2:A2)",
                "SUM(T!A2:A3)",
                "SUM(T!A3:A4)",
                "SUM(T!A3:A5)",
                "SUM(T!A2:A5)",
                "SUM(T!A4:A5)",
                "SUM(T!A:A)",
                "SUM(T!3:4)",
                "SUM(T!2:5)",
                "SUM(T!1:2)",
                "T!$A$5",
                "SUM(T!A5:B5)",
                "NameOne+1",
                "NameTwo",
                "SUM(#REF!)",
                "SUM(T!C1:C7)",
                "SUM(T!B:C)",
            ]),
            "{extension}"
        );
    });
}

#[test]
fn columns_deleted_give_what_libreoffice_gives() {
    let mods = WorkbookModifications {
        delete_columns: scopes(&["T!B"]),
        ..Default::default()
    };
    each(&mods, |extension, f| {
        assert_eq!(
            f[18..21],
            norm(&["SUM(T!B2:C2)", "SUM(T!B1:B10)", "SUM(T!B:B)"])[..],
            "{extension}"
        );
        assert_eq!(
            f[..4],
            norm(&["T!A2", "T!A4", "T!A5", "T!A6"])[..],
            "{extension}"
        );
    });
}

#[test]
fn a_sheet_deleted_leaves_ref_errors() {
    let mods = WorkbookModifications {
        delete_sheets: Some(HashSet::from(["T".to_string()])),
        ..Default::default()
    };
    each(&mods, |extension, f| {
        assert_eq!(f[0], "#REF!", "{extension}");
        assert_eq!(f[4], "SUM(#REF!)", "{extension}");
        // A name is not a reference: the formula keeps it, and the name now holds `#REF!`.
        assert_eq!(f[16..18], ["NameOne+1", "NameTwo"], "{extension}");
    });
}

#[test]
fn a_formula_that_reads_other_cells_has_no_cached_result_and_the_rest_keeps_it() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let output = dir.path().join(format!("out.{extension}"));
        let mods = WorkbookModifications {
            delete_rows: scopes(&["T!4:5"]),
            ..Default::default()
        };
        modify_workbook(&asset(extension), &output, &mods).unwrap();
        let book = Book::from_file(&output).unwrap().workbook().unwrap();
        let f = book.sheets.iter().find(|s| s.name == "F").unwrap();
        // `=$T.A2` does not change and `=$T.A6` (now `=$T.A4`) reads the same cell in another place:
        // both keep their result. `=$T.A4` reads a deleted cell and `SUM($T.A2:A4)` a range that
        // shrinks: they have none.
        let cached = |row: u32| f.cell(row, 0).unwrap().value.clone();
        let has_result = |row: u32| !matches!(cached(row), crate::reader::CellValue::Empty);
        assert!(has_result(0), "{extension}");
        assert!(has_result(3), "{extension}");
        assert!(!has_result(1), "{extension}");
        assert!(!has_result(5), "{extension}");
    }
}

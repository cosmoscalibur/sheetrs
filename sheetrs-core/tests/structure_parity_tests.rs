//! The same command on the XLSX and the ODS that LibreOffice wrote from one workbook
//! (`sheetrs-core/scripts/gen_structure_fixtures.py`) gives the same formulas. The expectations are
//! what LibreOffice 26.8 gave for the same deletion, asked through UNO.

use sheetrs_core::reader::read_workbook;
use sheetrs_core::writer::{ValueScope, WorkbookModifications, modify_workbook};
use std::collections::HashSet;
use std::path::{Path, PathBuf};

fn asset(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("tests/assets/structure/book.{extension}"))
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
    let book = read_workbook(&output).unwrap();
    let sheet = book.sheets.iter().find(|s| s.name == "F").unwrap();
    let mut cells: Vec<_> = sheet.cells.iter().filter(|(at, _)| at.1 == 0).collect();
    cells.sort_by_key(|(at, _)| at.0);
    cells
        .into_iter()
        .filter_map(|(_, cell)| cell.formula.as_deref())
        .map(|f| f.replace("T!#REF!", "#REF!"))
        .collect()
}

/// The formula with each range of one cell (`A1:A1`) as that cell: LibreOffice writes the 3D
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

fn both(mods: &WorkbookModifications) -> Vec<String> {
    let xlsx = run("xlsx", mods);
    let ods = run("ods", mods);
    let same = |formulas: &[String]| {
        formulas
            .iter()
            .map(|f| one_cell_ranges(f))
            .collect::<Vec<_>>()
    };
    assert_eq!(
        same(&xlsx),
        same(&ods),
        "the same command gives other formulas in each format"
    );
    xlsx
}

#[test]
fn rows_deleted_give_what_libreoffice_gives() {
    let f = both(&WorkbookModifications {
        delete_rows: scopes(&["T!4:5"]),
        ..Default::default()
    });
    // `Rows.removeByIndex(3, 2)` in LibreOffice; a reference that goes is `#REF!`.
    assert_eq!(
        f,
        [
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
            "SUM(T:V!A2:A3)",
            "SUM(T:V!A1)",
            "SUM(#REF!)",
            "SUM(T:V!A6:A8)",
            "SUM(T:V!B2:D2)",
            "SUM(U:V!A2:A6)",
            "SUM(T:U!A3:A7)",
            "SUM(#REF!)",
            "SUM(T:V!B1:B8)",
        ]
    );
}

#[test]
fn several_scopes_read_the_rows_of_the_input() {
    let f = both(&WorkbookModifications {
        // Rows 2 and 7-8 of the input, in any order.
        delete_rows: scopes(&["T!7:8", "T!2"]),
        ..Default::default()
    });
    assert_eq!(
        f,
        [
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
            "SUM(T:V!A2:A3)",
            "SUM(T:V!A1)",
            "SUM(T:V!A4:A5)",
            "SUM(T:V!A6:A6)",
            "SUM(#REF!)",
            "SUM(U:V!A2:A6)",
            "SUM(T:U!A3:A7)",
            "SUM(T:V!A4:B4)",
            "SUM(T:V!B1:B7)",
        ]
    );
}

#[test]
fn columns_deleted_give_the_same_in_both_formats() {
    let f = both(&WorkbookModifications {
        delete_columns: scopes(&["T!B"]),
        ..Default::default()
    });
    assert_eq!(f[18..21], ["SUM(T!B2:C2)", "SUM(T!B1:B10)", "SUM(T!B:B)"]);
    // Columns of a sheet of the span change the 3D references that name them.
    assert_eq!(
        f[21..],
        [
            "SUM(T:V!A2:A4)",
            "SUM(T:V!A1)",
            "SUM(T:V!A4:A5)",
            "SUM(T:V!A6:A8)",
            "SUM(T:V!B2:C2)",
            "SUM(U:V!A2:A6)",
            "SUM(T:U!A3:A9)",
            "SUM(T:V!A4:A4)",
            "SUM(#REF!)",
        ]
    );
    assert_eq!(f[..4], ["T!A2", "T!A4", "T!A5", "T!A6"]);
}

#[test]
fn a_sheet_removed_leaves_ref_errors_in_both_formats() {
    let f = both(&WorkbookModifications {
        remove_sheets: Some(HashSet::from(["T".to_string()])),
        ..Default::default()
    });
    assert_eq!(f[0], "#REF!");
    assert_eq!(f[4], "SUM(#REF!)");
    // A name is not a reference: the formula keeps it, and the name now holds `#REF!`.
    assert_eq!(f[16..18], ["NameOne+1", "NameTwo"]);
}

#[test]
fn a_formula_that_changes_has_no_cached_result_and_the_rest_keeps_it() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let output = dir.path().join(format!("out.{extension}"));
        let mods = WorkbookModifications {
            delete_rows: scopes(&["T!4:5"]),
            ..Default::default()
        };
        modify_workbook(&asset(extension), &output, &mods).unwrap();
        let book = read_workbook(&output).unwrap();
        let f = book.sheets.iter().find(|s| s.name == "F").unwrap();
        // `=$T.A2` does not change; `=$T.A6` does.
        let cached = |row: u32| f.get_cell(row, 0).unwrap().value.clone();
        assert!(
            !matches!(cached(0), sheetrs_core::reader::CellValue::Empty),
            "{extension}"
        );
        assert!(
            matches!(cached(3), sheetrs_core::reader::CellValue::Empty),
            "{extension}"
        );
    }
}

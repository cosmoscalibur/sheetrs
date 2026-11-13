//! Writing values into a workbook that serves as a template: the formulas stay, the cached results
//! are the new ones, the constants named are replaced. Both formats, the same operation.
//!
//! The template is `formulas.*` of `sheetrs-io/tests/assets` (made by `LibreOffice`): sheet `T` holds
//! `1` in `A1` and sheet `F` formulas such as `T!A1*2` in `A1` and `A2`.

use sheetrs_edit::{BookEdit, CachedResult, CellWrites, Constant};
use sheetrs_io::Book;
use sheetrs_io::reader::{CellValue, ExcelError, Workbook};
use std::path::{Path, PathBuf};

fn template(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-io/tests/assets")
        .join(format!("formulas.{extension}"))
}

fn cell(book: &Workbook, sheet: &str, row: u32, col: u32) -> sheetrs_io::reader::Cell {
    book.sheets
        .iter()
        .find(|s| s.name == sheet)
        .unwrap()
        .cells
        .get(&(row, col))
        .unwrap()
        .clone()
}

#[test]
fn results_and_constants_are_written_and_the_formulas_stay() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let before = Book::from_file(template(extension))
            .unwrap()
            .workbook()
            .unwrap();
        let formula_of = |book: &Workbook, row| cell(book, "F", row, 0).formula;

        let mut writes = CellWrites::new();
        writes
            .constant("T", 0, 0, Constant::Number(10.0))
            .result("F", 0, 0, CachedResult::Number(20.0))
            .result("F", 1, 0, CachedResult::Text("twenty & <more>".into()))
            .result("F", 2, 0, CachedResult::Boolean(true))
            .result("F", 3, 0, CachedResult::Error("#DIV/0!".into()));
        let mut editor = Book::from_file(template(extension)).unwrap();
        editor.write_cells(writes);
        assert!(
            !editor.explain().contains("recalculate"),
            "{extension}: the results written are taken as current"
        );
        let report = editor.write(&out).unwrap();
        assert_eq!(report.written, 5, "{extension}");

        let after = Book::from_file(&out).unwrap().workbook().unwrap();
        assert_eq!(
            cell(&after, "T", 0, 0).value,
            CellValue::Number(10.0),
            "{extension}"
        );
        assert_eq!(
            cell(&after, "F", 0, 0).value,
            CellValue::Number(20.0),
            "{extension}"
        );
        assert_eq!(
            cell(&after, "F", 1, 0).value,
            CellValue::Text("twenty & <more>".into()),
            "{extension}"
        );
        assert_eq!(
            cell(&after, "F", 2, 0).value,
            CellValue::Boolean(true),
            "{extension}"
        );
        assert_eq!(
            cell(&after, "F", 3, 0).value,
            CellValue::Error(ExcelError::DivZero),
            "{extension}"
        );
        // The formulas, whatever they were, are what they were.
        for row in 0..4 {
            assert_eq!(
                formula_of(&after, row),
                formula_of(&before, row),
                "{extension} row {row}"
            );
        }
        // The cells that were not written are not touched.
        assert_eq!(
            cell(&after, "F", 4, 0).value,
            cell(&before, "F", 4, 0).value,
            "{extension}"
        );
        assert_eq!(
            cell(&after, "F", 0, 0).num_fmt,
            cell(&before, "F", 0, 0).num_fmt,
            "{extension}: the style stays"
        );
    }
}

#[test]
fn a_constant_replaces_a_formula_and_empty_clears() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut writes = CellWrites::new();
        writes
            .constant("F", 0, 0, Constant::Number(7.0))
            .constant("F", 1, 0, Constant::Empty);
        let mut editor = Book::from_file(template(extension)).unwrap();
        editor.write_cells(writes);
        editor.write(&out).unwrap();
        let after = Book::from_file(&out).unwrap().workbook().unwrap();
        let seven = cell(&after, "F", 0, 0);
        assert_eq!(
            (seven.value, seven.formula),
            (CellValue::Number(7.0), None),
            "{extension}"
        );
        let cleared = after
            .sheets
            .iter()
            .find(|s| s.name == "F")
            .unwrap()
            .cells
            .get(&(1, 0))
            .cloned();
        assert!(
            cleared.is_none_or(|c| c.value == CellValue::Empty && c.formula.is_none()),
            "{extension}"
        );
    }
}

#[test]
fn a_write_is_a_pass_of_its_own_and_the_sheet_must_exist() {
    let source = Book::from_file(template("xlsx")).unwrap();
    let mut editor = source.clone();
    let mut writes = CellWrites::new();
    writes.result("Nope", 0, 0, CachedResult::Number(1.0));
    editor.write_cells(writes);
    assert!(editor.explain().contains("write 1 cells in Nope"));
    let dir = tempfile::tempdir().unwrap();
    assert!(editor.write(dir.path().join("o.xlsx")).is_err());
}

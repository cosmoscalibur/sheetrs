//! Editing in memory gives the workbook that editing files gives: the same passes, with the workbook
//! kept as bytes between them, so that it runs where there is no file system (a browser).

use sheetrs_edit::{BookEdit, ColumnRange, RowRange, ValueScope};
use sheetrs_io::Book;
use sheetrs_protocol::{Find, Replace};
use std::path::{Path, PathBuf};

/// The workbook of the structure oracle, which the engine calculates (the one of `book` has an
/// external reference, which it cannot).
fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/structure")
        .join(name)
}

/// A plan of several passes: a deletion, the results it left stale calculated again, a conversion
/// (which cannot share the pass of a change before it) and a replacement (a pass of its own); this is
/// about files and memory, what the recalculation gives is judged in `recalculation.rs`.
fn plan(editor: &mut Book) {
    editor
        .delete_rows(["T!2".parse::<RowRange>().unwrap()])
        .delete_columns(["U!B".parse::<ColumnRange>().unwrap()])
        .formula_to_value(["F".parse::<ValueScope>().unwrap()])
        .replace(
            [],
            [Replace {
                find: Find::literal("a"),
                with: "A".to_string(),
            }],
        );
}

fn same_bytes_as_files(extension: &str) {
    let path = asset(&format!("book.{extension}"));
    let from_file = Book::from_file(&path).unwrap();
    let mut editor = from_file.clone();
    plan(&mut editor);
    assert!(
        editor.explain().lines().count() > 3,
        "the plan has several passes"
    );

    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join(format!("out.{extension}"));
    let report = editor.write(&out).unwrap();
    let on_disk = std::fs::read(&out).unwrap();

    // From a file, to bytes.
    let mut bytes = Vec::new();
    let in_memory = editor.write_to(&mut bytes).unwrap();
    assert_eq!(bytes, on_disk, "a file as the source");
    assert_eq!(in_memory.replaced, report.replaced);

    // From bytes, to bytes: what a browser has.
    let from_bytes = Book::from_bytes(std::fs::read(&path).unwrap()).unwrap();
    let mut editor = from_bytes.clone();
    plan(&mut editor);
    let mut bytes = Vec::new();
    editor.write_to(&mut bytes).unwrap();
    assert_eq!(bytes, on_disk, "bytes as the source");
}

#[test]
fn xlsx_in_memory_is_the_workbook_the_files_give() {
    same_bytes_as_files("xlsx");
}

#[test]
fn ods_in_memory_is_the_workbook_the_files_give() {
    same_bytes_as_files("ods");
}

#[test]
fn an_empty_plan_is_an_error_in_memory_too() {
    let source = Book::from_file(asset("book.xlsx")).unwrap();
    assert!(source.clone().write_to(Vec::new()).is_err());
}

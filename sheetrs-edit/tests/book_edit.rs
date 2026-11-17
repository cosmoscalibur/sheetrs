//! The public API of `BookEdit`: the changes collected in a `Book`, `explain`, `write`.

use sheetrs_edit::{BookEdit, Error};
use sheetrs_io::Book;
use sheetrs_protocol::{Find, Replace};
use std::path::{Path, PathBuf};

fn book(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("tests/assets/structure/book.{extension}"))
}

fn sheet_names(path: &Path) -> Vec<String> {
    Book::from_file(path)
        .unwrap()
        .workbook()
        .unwrap()
        .sheets
        .into_iter()
        .map(|sheet| sheet.name)
        .collect()
}

#[test]
fn a_book_writes_the_changes_it_collected_for_its_source() {
    for extension in ["xlsx", "ods"] {
        let source = Book::from_file(book(extension)).unwrap();
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut editor = source.clone();
        editor
            .delete_rows(["T!3".parse().unwrap()])
            .delete_sheets(["V"]);
        assert!(!editor.operations().is_empty());
        assert!(!out.exists(), "nothing is written before save");
        editor.write(&out).unwrap();
        let names = sheet_names(&out);
        assert!(!names.contains(&"V".to_string()), "{extension}: {names:?}");
        // a book can be saved again, and the input is untouched.
        editor
            .write(dir.path().join(format!("again.{extension}")))
            .unwrap();
        assert!(sheet_names(&book(extension)).contains(&"V".to_string()));
    }
}

#[test]
fn a_book_edits_bytes_as_it_edits_a_file() {
    let path = book("xlsx");
    let from_file = Book::from_file(&path).unwrap();
    let from_bytes = Book::from_bytes(std::fs::read(&path).unwrap()).unwrap();
    let dir = tempfile::tempdir().unwrap();
    let edit = |source: &Book, name: &str| {
        let out = dir.path().join(name);
        source
            .clone()
            .delete_columns(["T!B".parse().unwrap()])
            .write(&out)
            .unwrap();
        sheet_names(&out)
    };
    assert_eq!(edit(&from_file, "a.xlsx"), edit(&from_bytes, "b.xlsx"));
}

#[test]
fn explain_says_what_write_will_do_without_touching_a_file() {
    let source = Book::from_file(book("xlsx"))
        .unwrap()
        .with_calculate_all(true);
    let mut editor = source.clone();
    editor
        .delete_rows(["T!2".parse().unwrap()])
        .formula_to_value(["F".parse().unwrap()]);
    let text = editor.explain();
    // The conversion is of the workbook as it is and the deletion is written against it, in one
    // pass; the formulas that were not converted are left with stale results, which the engine
    // calculates again as the last pass (and only if it cannot is the workbook marked).
    assert!(text.starts_with("pass 1 of 2:"), "{text}");
    assert!(text.find("formula to value").unwrap() < text.find("delete rows").unwrap());
    assert!(
        text.contains("pass 2 of 2:\n  recalculate every formula"),
        "{text}"
    );
    assert!(
        text.ends_with("mark it to be recalculated when opened\n"),
        "{text}"
    );
    // With `calculate_all` the conversion is of what the engine calculates; by default an XLSX whose results are right converts the
    // results the file caches, with no engine (an ODS always has the engine: see `calculation_mode.rs`).
    assert!(
        text.contains("calculated by the engine in the pass that writes them"),
        "{text}"
    );
    let mut cached = source.with_calculate_all(false);
    cached.formula_to_value(["F".parse().unwrap()]);
    assert!(
        !cached
            .explain()
            .contains("calculated by the engine in the pass that writes them")
    );
}

/// The order of the plan, as the explanation of the default mode says it (the explanation is the contract of what `write` will do).
#[test]
fn the_default_plan_is_replace_convert_and_delete_then_recalculate() {
    for extension in ["xlsx", "ods"] {
        let mut editor = Book::from_file(book(extension)).unwrap();
        editor
            .delete_rows(["T!2".parse().unwrap()])
            .formula_to_value(["F".parse().unwrap()])
            .replace(
                ["U".into()],
                [Replace {
                    find: Find::literal("9"),
                    with: "8".into(),
                }],
            );
        let text = editor.explain();
        assert!(
            text.contains("pass 1 of 3:") && text.contains("pass 3 of 3:"),
            "{extension}: {text}"
        );
        let replace = text.find("replace").unwrap();
        let convert = text.find("formula to value").unwrap();
        let delete = text.find("delete rows").unwrap();
        let recalculate = text.find("recalculate every formula").unwrap();
        assert!(
            replace < convert && convert < delete && delete < recalculate,
            "{extension}: {text}"
        );
        // The conversion and the deletion share the pass; the recalculation is said once.
        assert!(
            text.contains("formula to value in F\n  delete rows T!2:2"),
            "{extension}: {text}"
        );
        assert_eq!(text.matches("recalculate every formula").count(), 1);
    }
}

#[test]
fn a_book_without_changes_does_not_write_and_says_why_with_a_type() {
    let source = Book::from_file(book("xlsx")).unwrap();
    let dir = tempfile::tempdir().unwrap();
    let editor = source.clone();
    assert!(editor.operations().is_empty());
    assert!(matches!(
        editor.write(dir.path().join("out.xlsx")),
        Err(Error::EmptyPlan)
    ));
}

#[test]
fn write_in_place_replaces_the_file_and_a_failed_run_leaves_it_as_it_was() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join(format!("book.{extension}"));
        std::fs::copy(book(extension), &path).unwrap();
        let source = Book::from_file(&path).unwrap();

        // A run that fails: the sheet does not exist.
        let before = std::fs::read(&path).unwrap();
        let mut broken = source.clone();
        broken.delete_rows(["Nope!3".parse().unwrap()]);
        assert!(broken.write_in_place().is_err());
        assert_eq!(std::fs::read(&path).unwrap(), before, "{extension}");
        let leftovers: Vec<_> = std::fs::read_dir(dir.path())
            .unwrap()
            .map(|e| e.unwrap().file_name())
            .collect();
        assert_eq!(leftovers.len(), 1, "{extension}: {leftovers:?}");

        let mut editor = source.clone();
        editor.delete_sheets(["V"]);
        editor.write_in_place().unwrap();
        assert!(
            !sheet_names(&path).contains(&"V".to_string()),
            "{extension}"
        );
        assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
    }
}

#[test]
fn write_in_place_needs_a_file() {
    let bytes = Book::from_bytes(std::fs::read(book("xlsx")).unwrap()).unwrap();
    let mut editor = bytes.clone();
    editor.delete_sheets(["V"]);
    assert!(matches!(editor.write_in_place(), Err(Error::NoFile)));
}

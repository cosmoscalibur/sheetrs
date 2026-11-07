//! The public API of the editor: made for a `Source`, changes in order, `explain`, `save`.

use sheetrs_edit::{Editor, Error};
use sheetrs_io::reader::{Source, read_workbook};
use std::path::{Path, PathBuf};

fn book(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("tests/assets/structure/book.{extension}"))
}

fn sheet_names(path: &Path) -> Vec<String> {
    read_workbook(path)
        .unwrap()
        .sheets
        .into_iter()
        .map(|sheet| sheet.name)
        .collect()
}

#[test]
fn an_editor_saves_the_changes_it_collected_for_its_source() {
    for extension in ["xlsx", "ods"] {
        let source = Source::from_file(book(extension)).unwrap();
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut editor = Editor::new(&source);
        editor
            .delete_rows(["T!3".parse().unwrap()])
            .remove_sheets(["V"]);
        assert!(!editor.is_empty() && editor.marks_recalculation());
        assert!(!out.exists(), "nothing is written before save");
        editor.save(&out).unwrap();
        let names = sheet_names(&out);
        assert!(!names.contains(&"V".to_string()), "{extension}: {names:?}");
        // An editor can be saved again, and the input is untouched.
        editor
            .save(dir.path().join(format!("again.{extension}")))
            .unwrap();
        assert!(sheet_names(&book(extension)).contains(&"V".to_string()));
    }
}

#[test]
fn an_editor_edits_bytes_as_it_edits_a_file() {
    let path = book("xlsx");
    let from_file = Source::from_file(&path).unwrap();
    let from_bytes = Source::from_bytes(std::fs::read(&path).unwrap(), "xlsx".parse().unwrap());
    let dir = tempfile::tempdir().unwrap();
    let edit = |source: &Source, name: &str| {
        let out = dir.path().join(name);
        Editor::new(source)
            .delete_columns(["T!B".parse().unwrap()])
            .save(&out)
            .unwrap();
        sheet_names(&out)
    };
    assert_eq!(edit(&from_file, "a.xlsx"), edit(&from_bytes, "b.xlsx"));
}

#[test]
fn explain_says_what_save_will_do_without_touching_a_file() {
    let source = Source::from_file(book("xlsx")).unwrap();
    let mut editor = Editor::new(&source);
    editor
        .delete_rows(["T!2".parse().unwrap()])
        .convert_to_values(["F".parse().unwrap()]);
    let text = editor.explain().unwrap();
    assert!(text.starts_with("pass 1 of 2:"), "{text}");
    assert!(text.ends_with("recalculate when opened\n"), "{text}");
}

#[test]
fn an_editor_without_changes_does_not_save_and_says_why_with_a_type() {
    let source = Source::from_file(book("xlsx")).unwrap();
    let dir = tempfile::tempdir().unwrap();
    let editor = Editor::new(&source);
    assert!(editor.is_empty() && !editor.marks_recalculation());
    assert!(matches!(
        editor.save(dir.path().join("out.xlsx")),
        Err(Error::EmptyPlan)
    ));
}

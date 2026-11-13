//! `sheetrs query --list-sheets` and `--list-names`.

use std::path::Path;
use std::process::{Command, Output};

fn sheetrs(args: &[&str], file: &str) -> Output {
    let file = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-edit/tests/assets/structure")
        .join(file);
    Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .args(args)
        .arg("-i")
        .arg(file)
        .output()
        .unwrap()
}

fn json(out: &Output) -> serde_json::Value {
    assert!(
        out.status.success(),
        "{}",
        String::from_utf8_lossy(&out.stderr)
    );
    serde_json::from_slice(&out.stdout).unwrap()
}

#[test]
fn the_sheets_are_listed_in_order_in_both_formats() {
    for file in ["book.xlsx", "book.ods"] {
        let sheets = json(&sheetrs(&["query", "--list-sheets"], file));
        assert_eq!(sheets, serde_json::json!(["T", "U", "V", "F"]), "{file}");
    }
}

#[test]
fn the_defined_names_are_listed_with_what_they_refer_to() {
    for file in ["book.xlsx", "book.ods"] {
        let names = json(&sheetrs(&["query", "--list-names"], file));
        assert_eq!(names["NameOne"], "T!$A$6", "{file}");
        assert_eq!(names["NameTwo"], "T!$A$1:$A$10", "{file}");
        assert_eq!(names.as_object().unwrap().len(), 4, "{file}");
    }
}

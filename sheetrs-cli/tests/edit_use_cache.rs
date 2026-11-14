//! `sheetrs edit --use-cache`: the recalculation an edit leaves to do calculates only the formulas the edit
//! reaches, and the others keep the results the file caches, in an XLSX and in an ODS. The workbooks are
//! `tests/assets/use_cache/book.*` of `sheetrs-edit`: `Calc!A7` caches `from-excel`, a result the engine
//! would not give (it gives `t`), and nothing the deletion of `Data!3` reaches reads it.

use serde_json::Value;
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};
use std::process::Command;

fn fixture(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-edit/tests/assets/use_cache")
        .join(format!("book.{extension}"))
}

/// `Calc!A1:A9` as `A1 -> value`, after deleting `Data!3` with `flags`.
fn after_deleting(extension: &str, flags: &[&str]) -> BTreeMap<String, Value> {
    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join(format!("out.{extension}"));
    let edit = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .arg("edit")
        .arg("-i")
        .arg(fixture(extension))
        .args(["--delete-rows", "Data!3", "-o"])
        .arg(&out)
        .args(flags)
        .output()
        .unwrap();
    assert!(
        edit.status.success(),
        "{}",
        String::from_utf8_lossy(&edit.stderr)
    );
    let query = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .args(["query", "-i"])
        .arg(&out)
        .args(["--range", "Calc!A1:A9"])
        .output()
        .unwrap();
    assert!(
        query.status.success(),
        "{}",
        String::from_utf8_lossy(&query.stderr)
    );
    let answer: Value = serde_json::from_slice(&query.stdout).unwrap();
    answer["cells"]
        .as_array()
        .unwrap()
        .iter()
        .map(|cell| {
            (
                cell["cell"].as_str().unwrap().to_string(),
                cell["value"].clone(),
            )
        })
        .collect()
}

#[test]
fn the_flag_keeps_the_results_of_the_formulas_the_edit_does_not_reach() {
    for extension in ["xlsx", "ods"] {
        let kept = after_deleting(extension, &["--use-cache"]);
        let full = after_deleting(extension, &[]);
        // Nothing reads what the deletion removed: the file's result, where the engine gives another.
        assert_eq!(kept["A7"], "from-excel", "{extension}");
        assert_eq!(full["A7"], "t", "{extension}");
    }
}

#[test]
fn the_flag_gives_what_the_default_gives_for_the_formulas_the_edit_reaches() {
    for extension in ["xlsx", "ods"] {
        let kept = after_deleting(extension, &["--use-cache"]);
        let full = after_deleting(extension, &[]);
        // The range that shrank, the whole column the row leaves, `ROW()` and the name.
        for cell in ["A1", "A5", "A6", "A8"] {
            assert_eq!(kept[cell], full[cell], "{extension} {cell}");
        }
        assert_eq!(kept["A1"], 180.0, "{extension}");
    }
}

#[test]
fn the_default_calculates_every_formula() {
    for extension in ["xlsx", "ods"] {
        let full = after_deleting(extension, &[]);
        // `ROW()` of `Calc!A8`, which the file caches as 99.
        assert_eq!(full["A8"], 8.0, "{extension}");
        assert_eq!(full["A7"], "t", "{extension}");
    }
}

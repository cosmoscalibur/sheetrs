//! `sheetrs edit` and `--calculate-all`: by default (what `LibreOffice` does) the recalculation an edit leaves to do
//! calculates only the formulas the edit reaches, and the others keep the results the file caches; with
//! `--calculate-all` every formula is calculated again, in an XLSX and in an ODS. The workbooks are
//! `tests/assets/cached/book.*` of `sheetrs-edit`: `Calc!A7` caches `from-excel`, a result the engine
//! would not give (it gives `t`), and nothing the deletion of `Data!3` reaches reads it.

use serde_json::Value;
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};
use std::process::Command;

fn fixture(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-edit/tests/assets/cached")
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
fn by_default_the_results_of_the_formulas_the_edit_does_not_reach_are_kept() {
    for extension in ["xlsx", "ods"] {
        let kept = after_deleting(extension, &[]);
        let full = after_deleting(extension, &["--calculate-all"]);
        // Nothing reads what the deletion removed: the file's result, where the engine gives another.
        assert_eq!(kept["A7"], "from-excel", "{extension}");
        assert_eq!(full["A7"], "t", "{extension}");
    }
}

#[test]
fn by_default_the_formulas_the_edit_reaches_give_what_calculate_all_gives() {
    for extension in ["xlsx", "ods"] {
        let kept = after_deleting(extension, &[]);
        let full = after_deleting(extension, &["--calculate-all"]);
        // The range that shrank and the whole column the row leaves.
        for cell in ["A1", "A5"] {
            assert_eq!(kept[cell], full[cell], "{extension} {cell}");
        }
        assert_eq!(kept["A1"], 180.0, "{extension}");
        // `ROW()` of a sheet that lost no row and the name `rate` (`Data!$C$1`, above the row) are not reached, as in
        // `LibreOffice` (an XLSX saved and opened again keeps 99 and 140): they keep what the file caches; an ODS that caches
        // plain numbers in cells with no format has them calculated when `LibreOffice` opens it (14 and 8), which the default does too.
        let (rate, row) = if extension == "ods" {
            (14.0, 8.0)
        } else {
            (140.0, 99.0)
        };
        assert_eq!(kept["A6"], rate, "{extension}");
        assert_eq!(kept["A8"], row, "{extension}");
        assert_eq!(full["A6"], 14.0, "{extension}");
        assert_eq!(full["A8"], 8.0, "{extension}");
    }
}

#[test]
fn calculate_all_calculates_every_formula() {
    for extension in ["xlsx", "ods"] {
        let full = after_deleting(extension, &["--calculate-all"]);
        // `ROW()` of `Calc!A8`, which the file caches as 99.
        assert_eq!(full["A8"], 8.0, "{extension}");
        assert_eq!(full["A7"], "t", "{extension}");
    }
}

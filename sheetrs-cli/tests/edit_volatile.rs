//! `sheetrs edit --now --seed`: the recalculation an edit leaves to do runs with the clock and the seed
//! the flags give, so `NOW()`, `TODAY()` and `RAND()` of the result are the same on every run, in an
//! XLSX and in an ODS (`volatile.*` hold `NOW()`, `TODAY()` and `RAND()` in `V!A1:C1`).

use std::path::{Path, PathBuf};
use std::process::Command;

fn fixture(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-edit/tests/assets/volatile")
        .join(format!("volatile.{extension}"))
}

/// `V!A1:C1` after deleting row 3 with `flags`.
fn after_deleting(extension: &str, flags: &[&str]) -> Vec<f64> {
    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join(format!("out.{extension}"));
    let edit = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .arg("edit")
        .arg("-i")
        .arg(fixture(extension))
        .args(["--delete-rows", "V!3", "-o"])
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
        .args(["--range", "V!A1:C1"])
        .output()
        .unwrap();
    assert!(
        query.status.success(),
        "{}",
        String::from_utf8_lossy(&query.stderr)
    );
    let answer: serde_json::Value = serde_json::from_slice(&query.stdout).unwrap();
    answer["cells"]
        .as_array()
        .unwrap()
        .iter()
        .map(|cell| cell["value"].as_f64().unwrap())
        .collect()
}

#[test]
fn now_and_seed_fix_the_recalculation_an_edit_leaves() {
    for extension in ["xlsx", "ods"] {
        let flags = ["--now", "45366.75", "--seed", "7"];
        let first = after_deleting(extension, &flags);
        let again = after_deleting(extension, &flags);
        assert!((first[0] - 45366.75).abs() < 1e-9, "{extension}: {first:?}");
        assert!((first[1] - 45366.0).abs() < 1e-9, "{extension}: {first:?}");
        assert!((0.0..1.0).contains(&first[2]), "{extension}: {first:?}");
        assert_eq!(first, again, "{extension}");
        let other = after_deleting(extension, &["--seed", "8"]);
        assert!((first[2] - other[2]).abs() > 1e-12, "{extension}");
    }
}

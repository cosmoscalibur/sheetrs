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

fn config(text: &str) -> (tempfile::TempDir, String) {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("sheetrs.toml");
    std::fs::write(&path, text).unwrap();
    let path = path.to_str().unwrap().to_owned();
    (dir, path)
}

#[test]
fn now_and_seed_of_the_configuration_file_are_the_defaults_of_the_flags() {
    let (_dir, path) = config("[global]\nnow = 45366.75\nseed = 7\n");
    for extension in ["xlsx", "ods"] {
        let by_file = after_deleting(extension, &["--config", &path]);
        let by_flags = after_deleting(extension, &["--now", "45366.75", "--seed", "7"]);
        assert_eq!(by_file, by_flags, "{extension}");
        // A flag wins over the file, field by field.
        let seed = after_deleting(extension, &["--config", &path, "--seed", "8"]);
        assert!((seed[0] - 45366.75).abs() < 1e-9, "{extension}: {seed:?}");
        assert!((seed[2] - by_file[2]).abs() > 1e-12, "{extension}");
        let now = after_deleting(extension, &["--config", &path, "--now", "40000"]);
        assert!((now[0] - 40000.0).abs() < 1e-9, "{extension}: {now:?}");
        assert!((now[2] - by_file[2]).abs() < 1e-12, "{extension}");
    }
}

#[test]
fn a_seed_that_is_not_a_non_negative_integer_in_the_configuration_is_refused() {
    let (_dir, path) = config("[global]\nseed = -1\n");
    let dir = tempfile::tempdir().unwrap();
    let edit = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .args(["edit", "-i"])
        .arg(fixture("xlsx"))
        .args(["--delete-rows", "V!3", "--config", &path, "-o"])
        .arg(dir.path().join("out.xlsx"))
        .output()
        .unwrap();
    assert!(!edit.status.success());
    assert!(
        String::from_utf8_lossy(&edit.stderr).contains("seed"),
        "{}",
        String::from_utf8_lossy(&edit.stderr)
    );
}

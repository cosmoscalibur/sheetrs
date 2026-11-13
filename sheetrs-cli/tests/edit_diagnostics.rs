//! What `sheetrs edit` reports besides its result: a diagnostic on the standard error, as text, as a
//! JSON record or not at all, by the flags `--log-format` and `-q`.

use std::path::{Path, PathBuf};
use std::process::{Command, Output};

/// A workbook with a reference to another one: the engine cannot calculate it, so deleting rows from
/// it leaves it marked to be recalculated, and the run says so.
fn external() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../sheetrs-edit/tests/assets/book/book.xlsx")
}

fn delete_rows(extra: &[&str]) -> Output {
    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join("out.xlsx");
    let done = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .args(extra)
        .arg("edit")
        .arg("-i")
        .arg(external())
        .args(["--delete-rows", "Data!10:11", "-o"])
        .arg(&out)
        .output()
        .unwrap();
    assert!(
        done.status.success(),
        "{}",
        String::from_utf8_lossy(&done.stderr)
    );
    done
}

fn stderr(out: &Output) -> String {
    String::from_utf8_lossy(&out.stderr).into_owned()
}

#[test]
fn a_calculation_the_engine_cannot_do_is_a_warning_line_by_default() {
    let said = stderr(&delete_rows(&[]));
    assert!(
        said.starts_with("warning: the workbook was not calculated"),
        "{said}"
    );
}

#[test]
fn the_json_log_format_writes_one_record_with_the_stable_code_and_the_context() {
    let said = stderr(&delete_rows(&["--log-format", "json"]));
    let lines: Vec<&str> = said.lines().collect();
    assert_eq!(lines.len(), 1, "{said}");
    let record: serde_json::Value = serde_json::from_str(lines[0]).unwrap();
    assert_eq!(record["level"], "warning");
    assert_eq!(record["code"], "RECALC_NOT_SUPPORTED");
    assert!(
        record["message"]
            .as_str()
            .unwrap()
            .contains("will be done by the spreadsheet program")
    );
    assert!(!record["context"]["reason"].as_str().unwrap().is_empty());
    assert_eq!(record["context"]["operation"], "recalculate");
}

#[test]
fn quiet_keeps_only_the_errors_so_a_warning_is_not_written() {
    assert_eq!(stderr(&delete_rows(&["-q"])), "");
}

#[test]
fn quiet_and_verbose_cannot_be_given_together() {
    let done = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .args(["-q", "-v", "edit", "-i"])
        .arg(external())
        .output()
        .unwrap();
    assert!(!done.status.success());
}

#[test]
fn a_log_format_that_is_not_one_is_refused() {
    let done = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .args(["--log-format", "xml", "query", "-i"])
        .arg(external())
        .output()
        .unwrap();
    assert!(!done.status.success());
    assert!(String::from_utf8_lossy(&done.stderr).contains("human"));
}

//! What `sheetrs edit` does with the file it writes: never the original, and never a file that was
//! there before when the run fails.

use std::path::{Path, PathBuf};
use std::process::{Command, Output};

fn asset() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../sheetrs-edit/tests/assets/replace/input.xlsx")
}

fn edit(args: &[&str]) -> Output {
    Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .arg("edit")
        .args(args)
        .output()
        .unwrap()
}

fn text(out: &Output) -> String {
    String::from_utf8_lossy(&out.stderr).into_owned()
}

#[test]
fn an_output_that_is_the_input_is_refused_and_says_to_use_in_place() {
    let dir = tempfile::tempdir().unwrap();
    let book = dir.path().join("book.xlsx");
    std::fs::copy(asset(), &book).unwrap();
    let before = std::fs::read(&book).unwrap();
    let out = edit(&[
        "-i",
        book.to_str().unwrap(),
        "--delete-rows",
        "T!1",
        "-o",
        book.to_str().unwrap(),
    ]);
    assert!(!out.status.success());
    assert!(text(&out).contains("--in-place"), "{}", text(&out));
    assert_eq!(std::fs::read(&book).unwrap(), before);
}

#[test]
fn a_failed_run_leaves_an_output_that_was_there_as_it_was() {
    let dir = tempfile::tempdir().unwrap();
    let out_path = dir.path().join("already.xlsx");
    std::fs::write(&out_path, b"what the user had").unwrap();
    // A sheet the workbook does not have: the run fails.
    let out = edit(&[
        "-i",
        asset().to_str().unwrap(),
        "--delete-rows",
        "NoSuchSheet!1",
        "-o",
        out_path.to_str().unwrap(),
    ]);
    assert!(!out.status.success(), "{}", text(&out));
    assert_eq!(std::fs::read(&out_path).unwrap(), b"what the user had");
    // Nothing is left beside it.
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
}

#[test]
fn a_failed_run_removes_the_output_it_created() {
    let dir = tempfile::tempdir().unwrap();
    let out_path = dir.path().join("new.xlsx");
    let out = edit(&[
        "-i",
        asset().to_str().unwrap(),
        "--delete-rows",
        "NoSuchSheet!1",
        "-o",
        out_path.to_str().unwrap(),
    ]);
    assert!(!out.status.success());
    assert!(!out_path.exists());
}

/// A device is not a file the run created: it is written to and left where it was.
#[cfg(target_os = "linux")]
#[test]
fn a_failed_run_does_not_remove_a_device_given_as_the_output() {
    let device = Path::new("/dev/full");
    if !device.exists() {
        return;
    }
    let out = edit(&[
        "-i",
        asset().to_str().unwrap(),
        "--delete-rows",
        "T!1",
        "-o",
        "/dev/full",
    ]);
    assert!(!out.status.success(), "writing to /dev/full succeeded");
    assert!(device.exists(), "the device was removed");
    // The writer says what failed and nothing else (no line of the archive writer's own).
    assert!(!text(&out).contains("ZipWriter::drop"), "{}", text(&out));
}

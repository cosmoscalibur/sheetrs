//! `sheetrs edit --dry-run`: a preview that writes nothing and needs no output.

use std::path::{Path, PathBuf};
use std::process::{Command, Output};

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-edit/tests/assets/replace")
        .join(name)
}

fn edit(file: &Path, args: &[&str]) -> Output {
    Command::new(env!("CARGO_BIN_EXE_sheetrs"))
        .arg("edit")
        .arg("-i")
        .arg(file)
        .args(args)
        .output()
        .unwrap()
}

fn text(out: &Output) -> String {
    format!(
        "{}{}",
        String::from_utf8_lossy(&out.stdout),
        String::from_utf8_lossy(&out.stderr)
    )
}

/// A preview needs no output, shows the passes and leaves everything as it was.
#[test]
fn a_dry_run_needs_no_output_and_writes_nothing() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join(format!("book.{extension}"));
        std::fs::copy(asset(&format!("input.{extension}")), &path).unwrap();
        let before = std::fs::read(&path).unwrap();

        let dry = edit(&path, &["--delete-rows", "T!1", "--dry-run"]);
        let said = text(&dry);
        assert!(dry.status.success(), "{extension}: {said}");
        assert!(said.contains("[DRY RUN]"), "{said}");
        assert!(said.contains("delete rows"), "{said}");
        assert!(said.contains("Nothing is written"), "{said}");

        assert_eq!(
            std::fs::read(&path).unwrap(),
            before,
            "{extension}: the input changed"
        );
        let files = std::fs::read_dir(dir.path()).unwrap().count();
        assert_eq!(files, 1, "{extension}: a file appeared");
    }
}

/// With an output given, the preview says where the result would go and still writes nothing.
#[test]
fn a_dry_run_with_an_output_says_where_and_writes_nothing() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join(format!("book.{extension}"));
        std::fs::copy(asset(&format!("input.{extension}")), &path).unwrap();
        let before = std::fs::read(&path).unwrap();
        let out = dir.path().join(format!("out.{extension}"));

        let dry = edit(
            &path,
            &[
                "--delete-rows",
                "T!1",
                "--dry-run",
                "--output",
                out.to_str().unwrap(),
            ],
        );
        let said = text(&dry);
        assert!(dry.status.success(), "{extension}: {said}");
        assert!(said.contains("Output would be:"), "{said}");
        assert!(!out.exists(), "{extension}: the preview wrote the output");

        let in_place = edit(&path, &["--delete-rows", "T!1", "--dry-run", "--in-place"]);
        assert!(
            in_place.status.success(),
            "{extension}: {}",
            text(&in_place)
        );
        assert_eq!(
            std::fs::read(&path).unwrap(),
            before,
            "{extension}: --in-place changed the input"
        );
    }
}

/// The rest of the validation stays: a command that changes files still needs an output.
#[test]
fn a_command_that_writes_still_needs_an_output() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("book.xlsx");
    std::fs::copy(asset("input.xlsx"), &path).unwrap();
    let before = std::fs::read(&path).unwrap();

    let run = edit(&path, &["--delete-rows", "T!1"]);
    assert!(!run.status.success());
    assert!(
        text(&run).contains("Output file is required"),
        "{}",
        text(&run)
    );
    assert_eq!(std::fs::read(&path).unwrap(), before);
}

/// A sweep on the command line: every `--replace` has its `--with`, by position, each pair is an
/// operation in order; counts that do not match are an error that says so.
#[test]
fn repeated_replace_and_with_are_pairs_by_position() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join(format!("book.{extension}"));
        std::fs::copy(asset(&format!("input.{extension}")), &path).unwrap();

        let pairs = edit(
            &path,
            &[
                "--replace",
                "hello",
                "--with",
                "bye",
                "--replace",
                "bye",
                "--with",
                "gone",
                "--dry-run",
            ],
        );
        let said = text(&pairs);
        assert!(pairs.status.success(), "{extension}: {said}");
        assert!(said.matches("replace").count() >= 2, "{said}");

        let swept = dir.path().join(format!("swept.{extension}"));
        let swept_run = edit(
            &path,
            &[
                "--replace",
                "hello",
                "--with",
                "bye",
                "--replace",
                "bye",
                "--with",
                "gone",
                "-o",
                swept.to_str().unwrap(),
            ],
        );
        assert!(
            swept_run.status.success(),
            "{extension}: {}",
            text(&swept_run)
        );
        let first = dir.path().join(format!("first.{extension}"));
        let second = dir.path().join(format!("second.{extension}"));
        assert!(
            edit(
                &path,
                &[
                    "--replace",
                    "hello",
                    "--with",
                    "bye",
                    "-o",
                    first.to_str().unwrap()
                ]
            )
            .status
            .success()
        );
        assert!(
            edit(
                &first,
                &[
                    "--replace",
                    "bye",
                    "--with",
                    "gone",
                    "-o",
                    second.to_str().unwrap()
                ]
            )
            .status
            .success()
        );
        assert_eq!(
            std::fs::read(&swept).unwrap(),
            std::fs::read(&second).unwrap(),
            "{extension}: one command with two pairs is two commands"
        );

        let uneven = edit(
            &path,
            &[
                "--replace",
                "a",
                "--replace",
                "b",
                "--with",
                "c",
                "--dry-run",
            ],
        );
        let said = text(&uneven);
        assert!(!uneven.status.success(), "{said}");
        assert!(said.contains("2 --replace but 1 --with"), "{said}");
    }
}

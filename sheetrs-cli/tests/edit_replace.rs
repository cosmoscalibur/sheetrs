//! `sheetrs edit --replace`: the command says what it did, and a replacement and a conversion go in
//! one invocation.

use sheetrs_io::Book;
use sheetrs_io::reader::{Cell, CellValue};
use std::path::{Path, PathBuf};

/// The cell at `(row, col)` of a sheet, as the reader reads it back.
fn cell(path: &Path, sheet: &str, at: (u32, u32)) -> Cell {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let sheet = workbook.sheets.iter().find(|s| s.name == sheet).unwrap();
    sheet
        .cells
        .iter()
        .find(|((r, c), _)| (*r, *c) == at)
        .map(|(_, cell)| cell.clone())
        .unwrap()
}

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-edit/tests/assets/replace")
        .join(name)
}

#[test]
fn the_command_replaces_in_place_and_says_what_it_did() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join(format!("book.{extension}"));
        std::fs::copy(asset(&format!("input.{extension}")), &path).unwrap();
        let run = |args: &[&str]| {
            std::process::Command::new(env!("CARGO_BIN_EXE_sheetrs"))
                .arg("edit")
                .arg("-i")
                .arg(&path)
                .args(args)
                .output()
                .unwrap()
        };
        let dry = run(&[
            "--replace",
            "hello",
            "--with",
            "bye",
            "--scope",
            "U Sheet",
            "--dry-run",
            "--in-place",
        ]);
        let text = String::from_utf8_lossy(&dry.stdout).to_string();
        assert!(
            text.contains("replace \"hello\" with \"bye\" (literal, in value) in sheet U Sheet"),
            "{text}"
        );
        let done = run(&["--replace", "hello", "--with", "bye", "--in-place"]);
        let text = String::from_utf8_lossy(&done.stdout).to_string();
        assert!(
            done.status.success(),
            "{text}{}",
            String::from_utf8_lossy(&done.stderr)
        );
        assert!(text.contains("Replaced in 2 cells"), "{text}");
        assert_eq!(
            cell(&path, "T", (0, 0)).value,
            CellValue::Text("bye World bye".into()),
            "{extension}"
        );
        // `--with` without `--replace` is a usage error.
        assert!(!run(&["--with", "x", "--in-place"]).status.success());
    }
}

/// `sheetrs edit` has the engine: a replacement and a conversion of formulas to values go in one
/// invocation, and the conversion writes what the formulas give after the replacement (`F!A1` is
/// `T!A2`, which the replacement changes from 2 to 9), not the results the file cached.
#[test]
fn a_replacement_and_a_conversion_go_in_one_invocation() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let done = std::process::Command::new(env!("CARGO_BIN_EXE_sheetrs"))
            .arg("edit")
            .arg("-i")
            .arg(
                Path::new(env!("CARGO_MANIFEST_DIR"))
                    .join("../sheetrs-edit/tests/assets/structure")
                    .join(format!("book.{extension}")),
            )
            .args([
                "--replace",
                "2",
                "--with",
                "9",
                "--entire-cells",
                "--scope",
                "T!A2",
                "--formula-to-value",
                "F",
                "-o",
            ])
            .arg(&out)
            .output()
            .unwrap();
        assert!(
            done.status.success(),
            "{extension}: {}{}",
            String::from_utf8_lossy(&done.stdout),
            String::from_utf8_lossy(&done.stderr)
        );
        let found = cell(&out, "F", (0, 0));
        assert_eq!(found.value, CellValue::Number(9.0), "{extension}");
        assert!(found.formula.is_none(), "{extension}: no formula is left");
    }
}

/// An empty search matches nothing, as in `LibreOffice` (`replaceAll` changes no cell): the command
/// reports no cell and leaves every text as it was, in every mode.
#[test]
fn the_empty_search_replaces_nothing_in_every_mode() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join(format!("book.{extension}"));
        std::fs::copy(asset(&format!("input.{extension}")), &path).unwrap();
        for mode in [None, Some("--regex"), Some("--wildcards")] {
            for entire in [None, Some("--entire-cells")] {
                let out = std::process::Command::new(env!("CARGO_BIN_EXE_sheetrs"))
                    .arg("edit")
                    .arg("-i")
                    .arg(&path)
                    .args(["--replace", "", "--with", "Z", "--in-place"])
                    .args(mode)
                    .args(entire)
                    .output()
                    .unwrap();
                let text = String::from_utf8_lossy(&out.stdout).to_string();
                assert!(
                    out.status.success(),
                    "{text}{}",
                    String::from_utf8_lossy(&out.stderr)
                );
                assert!(text.contains("Replaced in 0 cells"), "{text}");
                assert!(!text.contains("not replaced"), "{text}");
                assert_eq!(
                    cell(&path, "T", (0, 0)).value,
                    CellValue::Text("Hello World hello".into()),
                    "{extension} {mode:?} {entire:?}"
                );
            }
        }
    }
}

//! A workbook that is its own template: `sheetc run --embed-template` and an executable built with
//! `--embed-template`. They give no JSON: the answer is a copy of the workbook with the result of
//! every formula and the inputs given in their cells, in the formats of its cells. The outputs the
//! workbook declares are discarded, its inputs are what is accepted, and a workbook that declares
//! none is recalculated as it is only if `--recalculate` says so.
//!
//! The workbooks were made by `LibreOffice` in both formats (`sheetrs-io/tests/assets`): `T!A1` holds
//! `1`, and sheet `F` formulas such as `T!A1*2`.

use sheetrs_edit::{CellWrites, Constant, Editor};
use sheetrs_io::reader::{CellValue, Source, read_workbook};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};

const SHEETC: &str = env!("CARGO_BIN_EXE_sheetc");

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-io/tests/assets")
        .join(name)
}

fn run(cmd: &str, args: &[&str], stdin: &str) -> (i32, String, String) {
    let mut child = Command::new(cmd)
        .args(args)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    child
        .stdin
        .take()
        .unwrap()
        .write_all(stdin.as_bytes())
        .unwrap();
    let out = child.wait_with_output().unwrap();
    (
        out.status.code().unwrap_or(-1),
        String::from_utf8_lossy(&out.stdout).into_owned(),
        String::from_utf8_lossy(&out.stderr).into_owned(),
    )
}

fn value(path: &Path, sheet: &str, row: u32, col: u32) -> CellValue {
    read_workbook(path)
        .unwrap()
        .sheets
        .into_iter()
        .find(|s| s.name == sheet)
        .unwrap()
        .cells
        .get(&(row, col))
        .unwrap()
        .value
        .clone()
}

fn settings(dir: &Path, json: &str) -> String {
    let path = dir.join("settings.json");
    std::fs::write(&path, json).unwrap();
    path.to_str().unwrap().to_owned()
}

const INPUT: &str = r#"{"inputs": {"base": "T!A1"}}"#;

fn path(p: &Path) -> &str {
    p.to_str().unwrap()
}

#[test]
fn run_writes_the_workbook_with_the_inputs_and_the_results() {
    for ext in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("answer.{ext}"));
        let template = asset(&format!("formulas.{ext}"));
        let (code, stdout, err) = run(
            SHEETC,
            &[
                "run",
                path(&template),
                "--embed-template",
                "--settings-file",
                &settings(dir.path(), INPUT),
                "--output",
                path(&out),
            ],
            r#"{"inputs": {"base": 10}}"#,
        );
        assert_eq!(code, 0, "{ext}: {err}");
        assert!(
            stdout.is_empty(),
            "{ext}: the answer is the workbook, not JSON"
        );
        assert_eq!(value(&out, "T", 0, 0), CellValue::Number(10.0), "{ext}");
        assert_eq!(value(&out, "F", 0, 0), CellValue::Number(20.0), "{ext}");
        assert_eq!(value(&out, "F", 3, 0), CellValue::Number(11.0), "{ext}");
        assert_eq!(
            value(&template, "T", 0, 0),
            CellValue::Number(1.0),
            "the template is untouched"
        );
    }
}

#[test]
fn the_outputs_a_workbook_declares_are_discarded() {
    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join("answer.xlsx");
    let (code, _, err) = run(
        SHEETC,
        &[
            "run",
            path(&asset("formulas.xlsx")),
            "--embed-template",
            "--settings-file",
            &settings(
                dir.path(),
                r#"{"inputs": {"base": "T!A1"}, "outputs": {"first": "F!A1"}}"#,
            ),
            "--output",
            path(&out),
        ],
        r#"{"inputs": {"base": 3}}"#,
    );
    assert_eq!(code, 0, "{err}");
    assert!(
        err.contains("outputs the workbook declares are discarded"),
        "{err}"
    );
    // Not only `F!A1`: all of the workbook is written.
    assert_eq!(value(&out, "F", 3, 0), CellValue::Number(4.0));
}

#[test]
fn a_workbook_without_inputs_is_recalculated_only_if_asked() {
    for ext in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        // A template whose cached results are stale: its input changed and its formulas did not.
        let stale = dir.path().join(format!("stale.{ext}"));
        let mut writes = CellWrites::new();
        writes.constant("T", 0, 0, Constant::Number(10.0));
        Editor::new(&Source::from_file(asset(&format!("formulas.{ext}"))).unwrap())
            .write_cells(writes)
            .save(&stale)
            .unwrap();
        assert_eq!(value(&stale, "F", 0, 0), CellValue::Number(2.0), "{ext}");

        let out = dir.path().join(format!("recalculated.{ext}"));
        let (code, _, err) = run(
            SHEETC,
            &[
                "run",
                path(&stale),
                "--embed-template",
                "--output",
                path(&out),
            ],
            "",
        );
        assert_eq!(
            code, 2,
            "{ext}: without --recalculate it is an error: {err}"
        );
        assert!(err.contains("--recalculate"), "{ext}: {err}");

        let (code, _, err) = run(
            SHEETC,
            &[
                "run",
                path(&stale),
                "--embed-template",
                "--recalculate",
                "--output",
                path(&out),
            ],
            "",
        );
        assert_eq!(code, 0, "{ext}: {err}");
        assert_eq!(value(&out, "F", 0, 0), CellValue::Number(20.0), "{ext}");
    }
}

#[test]
fn the_options_say_what_they_are_for() {
    let dir = tempfile::tempdir().unwrap();
    let template = asset("formulas.xlsx");
    let out = dir.path().join("o.xlsx");
    let with_inputs = settings(dir.path(), INPUT);
    let cases: [(&[&str], &str); 4] = [
        (
            &[
                "--embed-template",
                "--recalculate",
                "--settings-file",
                &with_inputs,
                "--output",
                path(&out),
            ],
            "declares some",
        ),
        (
            &["--recalculate", "--output", path(&out)],
            "need --embed-template",
        ),
        (&["--output", path(&out)], "need --embed-template"),
        (
            &["--embed-template", "--settings-file", &with_inputs],
            "give --output",
        ),
    ];
    for (flags, wanted) in cases {
        let mut args = vec!["run", path(&template)];
        args.extend_from_slice(flags);
        let (code, _, err) = run(SHEETC, &args, "{}");
        assert_eq!(code, 2, "{flags:?}: {err}");
        assert!(err.contains(wanted), "{flags:?}: {err}");
    }
}

#[test]
fn an_executable_carries_its_workbook_and_gives_only_the_workbook() {
    for ext in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let exe = dir.path().join("model");
        let template = asset(&format!("formulas.{ext}"));
        let (code, _, err) = run(
            SHEETC,
            &[
                "build",
                path(&template),
                "--embed-template",
                "--settings-file",
                &settings(dir.path(), INPUT),
                "-o",
                path(&exe),
            ],
            "",
        );
        assert_eq!(code, 0, "{ext}: {err}");

        let out = dir.path().join(format!("a.{ext}"));
        let (code, _, err) = run(
            path(&exe),
            &["--output", path(&out)],
            r#"{"inputs": {"base": 5}}"#,
        );
        assert_eq!(code, 0, "{ext}: {err}");
        assert_eq!(value(&out, "F", 0, 0), CellValue::Number(10.0), "{ext}");

        // No JSON: the ranges are not asked for, and the output is required.
        let (code, _, err) = run(path(&exe), &["--range", "F!A1"], "{}");
        assert_eq!(code, 2, "{ext}");
        assert!(err.contains("writes a workbook, not JSON"), "{ext}: {err}");
        // The metadata is answered as in any executable.
        let (code, meta, _) = run(path(&exe), &["--meta"], "");
        assert_eq!(code, 0, "{ext}");
        assert!(meta.contains("\"base\""), "{ext}: {meta}");
    }
}

#[test]
fn an_executable_without_its_workbook_does_not_write_one() {
    let dir = tempfile::tempdir().unwrap();
    let exe = dir.path().join("json");
    let template = asset("formulas.xlsx");
    let settings = dir.path().join("settings.json");
    std::fs::write(
        &settings,
        r#"{"inputs": {"base": "T!A1"}, "outputs": {"first": "F!A1"}}"#,
    )
    .unwrap();
    let (code, _, err) = run(
        SHEETC,
        &[
            "build",
            path(&template),
            "--settings-file",
            path(&settings),
            "--runtime",
            SHEETC,
            "-o",
            path(&exe),
        ],
        "",
    );
    assert_eq!(code, 0, "{err}");
    let (code, _, err) = run(
        path(&exe),
        &["--output", "x.xlsx"],
        r#"{"inputs": {"base": 1}}"#,
    );
    assert_eq!(code, 2);
    assert!(err.contains("carries no workbook"), "{err}");
    let (code, out, _) = run(path(&exe), &[], r#"{"inputs": {"base": 4}}"#);
    assert_eq!(code, 0);
    assert!(out.contains("\"first\""), "{out}");
}

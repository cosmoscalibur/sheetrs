//! The answer written into a workbook: `sheetc run --write-workbook` and an executable that carries
//! its template (`build --embed-template`). The output is the template with the results of the run
//! in every formula and the inputs given in their cells; the formulas and formats stay.
//!
//! The workbooks were made by `LibreOffice` in both formats (`sheetrs-io/tests/assets`): `T!A1` holds
//! `1`, and sheet `F` formulas such as `T!A1*2`.

use sheetrs_io::reader::{CellValue, read_workbook};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};

const SHEETC: &str = env!("CARGO_BIN_EXE_sheetc");

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-io/tests/assets")
        .join(name)
}

fn run(cmd: &str, args: &[&str], stdin: &str) -> (i32, String) {
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

fn settings(dir: &Path) -> String {
    let path = dir.join("settings.json");
    std::fs::write(&path, r#"{"inputs": {"base": "T!A1"}}"#).unwrap();
    path.to_str().unwrap().to_owned()
}

#[test]
fn run_writes_the_answer_into_a_copy_of_the_workbook() {
    for ext in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("answer.{ext}"));
        let template = asset(&format!("formulas.{ext}"));
        let (code, err) = run(
            SHEETC,
            &[
                "run",
                template.to_str().unwrap(),
                "--outputs",
                "any",
                "--settings-file",
                &settings(dir.path()),
                "--write-workbook",
                out.to_str().unwrap(),
            ],
            r#"{"inputs": {"base": 10}}"#,
        );
        assert_eq!(code, 0, "{ext}: {err}");
        // The input is where it belongs and the formulas hold the results of this run.
        assert_eq!(value(&out, "T", 0, 0), CellValue::Number(10.0), "{ext}");
        assert_eq!(value(&out, "F", 0, 0), CellValue::Number(20.0), "{ext}");
        assert_eq!(value(&out, "F", 3, 0), CellValue::Number(11.0), "{ext}");
        // The template is untouched.
        assert_eq!(value(&template, "T", 0, 0), CellValue::Number(1.0), "{ext}");
    }
}

#[test]
fn a_program_that_prunes_cannot_fill_a_template() {
    let dir = tempfile::tempdir().unwrap();
    let settings = dir.path().join("settings.json");
    std::fs::write(
        &settings,
        r#"{"inputs": {"base": "T!A1"}, "outputs": {"first": "F!A1"}}"#,
    )
    .unwrap();
    let template = asset("formulas.xlsx");
    let (code, err) = run(
        SHEETC,
        &[
            "run",
            template.to_str().unwrap(),
            "--settings-file",
            settings.to_str().unwrap(),
            "--write-workbook",
            dir.path().join("o.xlsx").to_str().unwrap(),
        ],
        "{}",
    );
    assert_eq!(code, 1);
    assert!(err.contains("outputs = any"), "{err}");
}

#[test]
fn an_executable_carries_its_template_and_an_external_one_replaces_it() {
    for ext in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let settings = settings(dir.path());
        let exe = dir.path().join("model");
        let template = asset(&format!("formulas.{ext}"));
        let (code, err) = run(
            SHEETC,
            &[
                "build",
                template.to_str().unwrap(),
                "--outputs",
                "any",
                "--settings-file",
                &settings,
                "--embed-template",
                "-o",
                exe.to_str().unwrap(),
            ],
            "",
        );
        assert_eq!(code, 0, "{ext}: {err}");

        let out = dir.path().join(format!("a.{ext}"));
        let (code, err) = run(
            exe.to_str().unwrap(),
            &["--write-workbook", out.to_str().unwrap()],
            r#"{"inputs": {"base": 5}}"#,
        );
        assert_eq!(code, 0, "{ext}: {err}");
        assert_eq!(value(&out, "F", 0, 0), CellValue::Number(10.0), "{ext}");

        // The same program, written into the other format's template of the same layout.
        let other = if ext == "xlsx" { "ods" } else { "xlsx" };
        let external = asset(&format!("formulas.{other}"));
        let out2 = dir.path().join(format!("b.{other}"));
        let (code, err) = run(
            exe.to_str().unwrap(),
            &[
                "--write-workbook",
                out2.to_str().unwrap(),
                "--template",
                external.to_str().unwrap(),
            ],
            r#"{"inputs": {"base": 7}}"#,
        );
        assert_eq!(code, 0, "{ext}: {err}");
        assert_eq!(value(&out2, "F", 0, 0), CellValue::Number(14.0), "{ext}");

        // The executable still answers JSON as any other.
        let (code, _) = run(
            exe.to_str().unwrap(),
            &["--range", "F!A1"],
            r#"{"inputs": {"base": 1}}"#,
        );
        assert_eq!(code, 0, "{ext}");
    }
}

#[test]
fn an_executable_without_a_template_says_so() {
    let dir = tempfile::tempdir().unwrap();
    let exe = dir.path().join("model");
    let template = asset("formulas.xlsx");
    let (code, err) = run(
        SHEETC,
        &[
            "build",
            template.to_str().unwrap(),
            "--outputs",
            "any",
            "--settings-file",
            &settings(dir.path()),
            "--runtime",
            SHEETC,
            "-o",
            exe.to_str().unwrap(),
        ],
        "",
    );
    assert_eq!(code, 0, "{err}");
    let (code, err) = run(
        exe.to_str().unwrap(),
        &[
            "--write-workbook",
            dir.path().join("o.xlsx").to_str().unwrap(),
        ],
        "{}",
    );
    assert_eq!(code, 2);
    assert!(err.contains("carries no template"), "{err}");
}

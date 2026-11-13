//! A workbook that is its own template: `sheetc calculate -o OUT` (the workbook is there, so it is the
//! template of its own answer) and an executable built with `--embed-template` (which carries it).
//! `-o OUT` gives a copy of the workbook with the result of every formula and the inputs given in
//! their cells, in the formats of its cells; without it the answer is JSON, as in any executable, and
//! `-o OUT --json` gives both from one calculation. With nothing declared the formula cells are the
//! outputs (the constants are already in the template) and any constant an input (a range, or a
//! defined name inside them); what the settings declare narrows them, as in any program, and the
//! formulas outside them keep the result the template stored. Inputs and outputs can always be given
//! (`--request`, `--range`), and with nothing given every formula is calculated again.
//!
//! The workbooks were made by `LibreOffice` in both formats (`sheetrs-io/tests/assets`): `T!A1` holds
//! `1`, and sheet `F` formulas such as `T!A1*2`.

use sheetrs_io::Book;
use sheetrs_io::reader::CellValue;
use sheetrs_io::write::{CellWrites, Constant, WorkbookModifications};
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
    // The request is read from the standard input only when it is asked for: `--request -`.
    let mut args = args.to_vec();
    if !stdin.is_empty() && !args.contains(&"--serve") && !args.contains(&"--request") {
        args.extend(["--request", "-"]);
    }
    let mut child = Command::new(cmd)
        .args(&args)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    // A command that refuses its arguments exits without reading the request: writing to it then
    // fails with a broken pipe, which is not what the test is about (and made it fail now and then).
    let _ = child.stdin.take().unwrap().write_all(stdin.as_bytes());
    let out = child.wait_with_output().unwrap();
    (
        out.status.code().unwrap_or(-1),
        String::from_utf8_lossy(&out.stdout).into_owned(),
        String::from_utf8_lossy(&out.stderr).into_owned(),
    )
}

fn value(path: &Path, sheet: &str, row: u32, col: u32) -> CellValue {
    Book::from_file(path)
        .unwrap()
        .workbook()
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

const INPUT: &str = r#"{"inputs": [{"name": "base", "range": "T!A1"}]}"#;

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
            &["calculate", "-i", path(&template), "--output", path(&out)],
            // Any constant is an input, by its range.
            r#"{"inputs": {"T!A1": 10}}"#,
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
        // What is not an input is refused, and nothing is written.
        let refused = dir.path().join(format!("refused.{ext}"));
        let (code, _, err) = run(
            SHEETC,
            &[
                "calculate",
                "-i",
                path(&template),
                "--output",
                path(&refused),
            ],
            r#"{"inputs": {"F!A1": 10}}"#,
        );
        assert_eq!(code, 1, "{ext}");
        assert!(err.contains("not an input"), "{ext}: {err}");
        assert!(!refused.exists(), "{ext}");
    }
}

#[test]
fn an_input_range_is_written_as_given_though_the_constants_are_not_outputs() {
    // The outputs of an embedded workbook are its formulas: an input is written from what the
    // request gave, not read back from the answer.
    for ext in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("answer.{ext}"));
        let template = asset(&format!("formulas.{ext}"));
        let (code, _, err) = run(
            SHEETC,
            &["calculate", "-i", path(&template), "--output", path(&out)],
            r#"{"inputs": {"T!A1:A3": [5, 6, 7], "T!B1:B2": [[1], [2]]}}"#,
        );
        assert_eq!(code, 0, "{ext}: {err}");
        for (row, wanted) in [(0, 5.0), (1, 6.0), (2, 7.0)] {
            assert_eq!(value(&out, "T", row, 0), CellValue::Number(wanted), "{ext}");
        }
        assert_eq!(value(&out, "T", 0, 1), CellValue::Number(1.0), "{ext}");
        assert_eq!(value(&out, "T", 1, 1), CellValue::Number(2.0), "{ext}");
        assert_eq!(
            value(&out, "T", 2, 1),
            CellValue::Number(30.0),
            "{ext}: not given"
        );
        // The formulas that read them are the answer.
        assert_eq!(value(&out, "F", 0, 0), CellValue::Number(10.0), "{ext}");
        assert_eq!(value(&out, "F", 4, 0), CellValue::Number(18.0), "{ext}");
    }
}

#[test]
fn what_the_settings_declare_narrows_the_json_the_inputs_and_what_the_workbook_writes() {
    let dir = tempfile::tempdir().unwrap();
    let both = r#"{"inputs": [{"name": "base", "range": "T!A1"}], "outputs": [{"name": "first", "range": "F!A1"}]}"#;
    let template = asset("formulas.xlsx");
    let out = dir.path().join("answer.xlsx");
    // The declared input is accepted by its name; the declared output is written, and the formulas
    // outside the program keep the result the template stored.
    let (code, stdout, err) = run(
        SHEETC,
        &[
            "calculate",
            "-i",
            path(&template),
            "--settings-file",
            &settings(dir.path(), both),
            "--output",
            path(&out),
        ],
        r#"{"inputs": {"base": 3}}"#,
    );
    assert_eq!(code, 0, "{err}");
    assert!(stdout.is_empty(), "the workbook alone gives no JSON");
    assert!(!err.contains("discarded"), "{err}");
    assert_eq!(value(&out, "T", 0, 0), CellValue::Number(3.0));
    assert_eq!(value(&out, "F", 0, 0), CellValue::Number(6.0));
    let stored = value(&template, "F", 3, 0);
    assert_eq!(
        value(&out, "F", 3, 0),
        stored,
        "F!A4 is not an output: it keeps its stored result"
    );
    // The JSON answer gives what was declared and nothing else, from the same calculation.
    let both_out = dir.path().join("both.xlsx");
    let (code, stdout, err) = run(
        SHEETC,
        &[
            "calculate",
            "-i",
            path(&template),
            "--settings-file",
            &settings(dir.path(), both),
            "--output",
            path(&both_out),
            "--json",
        ],
        r#"{"inputs": {"base": 3}}"#,
    );
    assert_eq!(code, 0, "{err}");
    let json: serde_json::Value = serde_json::from_str(&stdout).unwrap();
    assert_eq!(
        json,
        serde_json::json!({"valueRanges": [{"name": "first", "range": "F!A1", "values": [[6]]}]})
    );
    assert_eq!(value(&both_out, "F", 0, 0), CellValue::Number(6.0));
    assert_eq!(value(&both_out, "F", 3, 0), stored);
    // The same without `--output`: the JSON alone.
    let (code, stdout, err) = run(
        SHEETC,
        &[
            "calculate",
            "-i",
            path(&template),
            "--settings-file",
            &settings(dir.path(), both),
        ],
        r#"{"inputs": {"base": 3}}"#,
    );
    assert_eq!(code, 0, "{err}");
    let json: serde_json::Value = serde_json::from_str(&stdout).unwrap();
    assert_eq!(
        json,
        serde_json::json!({"valueRanges": [{"name": "first", "range": "F!A1", "values": [[6]]}]})
    );
    // Declaring the inputs narrows them: what is not declared is not an input, and nothing is written.
    let refused = dir.path().join("refused.xlsx");
    let (code, _, err) = run(
        SHEETC,
        &[
            "calculate",
            "-i",
            path(&template),
            "--settings-file",
            &settings(dir.path(), INPUT),
            "--output",
            path(&refused),
        ],
        r#"{"inputs": {"T!B1": 3}}"#,
    );
    assert_eq!(code, 1);
    assert!(err.contains("not an input"), "{err}");
    assert!(!refused.exists());
    // What a request may read is what was declared.
    let (code, _, err) = run(
        SHEETC,
        &[
            "calculate",
            "-i",
            path(&template),
            "--settings-file",
            &settings(dir.path(), both),
            "--range",
            "F!A4",
        ],
        "{}",
    );
    assert_eq!(code, 1);
    assert!(
        err.contains("outside the outputs this program declares"),
        "{err}"
    );
}

#[test]
fn without_outputs_declared_the_json_gives_the_formula_cells() {
    for ext in ["xlsx", "ods"] {
        let template = asset(&format!("formulas.{ext}"));
        let (code, stdout, err) = run(
            SHEETC,
            &["calculate", "-i", path(&template), "--range", "F!A1"],
            r#"{"inputs": {"T!A1": 10}}"#,
        );
        assert_eq!(code, 0, "{ext}: {err}");
        let json: serde_json::Value = serde_json::from_str(&stdout).unwrap();
        assert_eq!(
            json["valueRanges"][0]["values"],
            serde_json::json!([[20]]),
            "{ext}"
        );
        // A constant is not an output: it is in the template already.
        let (code, _, err) = run(
            SHEETC,
            &["calculate", "-i", path(&template), "--range", "T!A1"],
            "{}",
        );
        assert_eq!(code, 1, "{ext}");
        assert!(err.contains("not inside the outputs"), "{ext}: {err}");
    }
}

#[test]
fn with_nothing_given_every_formula_is_calculated_again_and_inputs_and_outputs_are_optional() {
    for ext in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        // A workbook whose cached results are stale: its input changed and its formulas did not.
        let stale = dir.path().join(format!("stale.{ext}"));
        let mut writes = CellWrites::new();
        writes.constant("T", 0, 0, Constant::Number(10.0));
        let mods = WorkbookModifications {
            replacements: writes.sheets().clone(),
            ..WorkbookModifications::default()
        };
        Book::from_file(asset(&format!("formulas.{ext}")))
            .unwrap()
            .modify(&stale, &mods)
            .unwrap();
        assert_eq!(value(&stale, "F", 0, 0), CellValue::Number(2.0), "{ext}");

        // Nothing declared and nothing given: the workbook as it is, every formula calculated again.
        let out = dir.path().join(format!("recalculated.{ext}"));
        let (code, _, err) = run(
            SHEETC,
            &["calculate", "-i", path(&stale), "--output", path(&out)],
            "",
        );
        assert_eq!(code, 0, "{ext}: {err}");
        assert_eq!(value(&out, "F", 0, 0), CellValue::Number(20.0), "{ext}");

        // Inputs can always be given: the workbook with them written and the results calculated.
        let given = dir.path().join(format!("given.{ext}"));
        let (code, _, err) = run(
            SHEETC,
            &["calculate", "-i", path(&stale), "--output", path(&given)],
            r#"{"inputs": {"T!A1": 5}}"#,
        );
        assert_eq!(code, 0, "{ext}: {err}");
        assert_eq!(value(&given, "T", 0, 0), CellValue::Number(5.0), "{ext}");
        assert_eq!(value(&given, "F", 0, 0), CellValue::Number(10.0), "{ext}");

        // What the settings declare is honored, not ignored: the formulas outside the outputs keep
        // the result the template stored.
        let declared = dir.path().join(format!("declared.{ext}"));
        let (code, _, err) = run(
            SHEETC,
            &[
                "calculate",
                "-i",
                path(&stale),
                "--settings-file",
                &settings(
                    dir.path(),
                    r#"{"inputs": [{"name": "base", "range": "T!A1"}], "outputs": [{"name": "first", "range": "F!A1"}]}"#,
                ),
                "--output",
                path(&declared),
            ],
            "",
        );
        assert_eq!(code, 0, "{ext}: {err}");
        assert_eq!(
            value(&declared, "F", 0, 0),
            CellValue::Number(20.0),
            "{ext}"
        );
        assert_eq!(
            value(&declared, "F", 3, 0),
            value(&stale, "F", 3, 0),
            "{ext}: F!A4 is not an output: it keeps its stored result"
        );
    }
}

/// A calculation that is given no `--request` must not wait for the standard input: it read the
/// standard input to the end after the whole calculation, and a pipe that nobody closes (a harness, a
/// pipeline whose first command was still running) held the process for ever (found with a workbook
/// of 6.4 M formulas, 23 s of CPU and then a wait on a futex: the main thread was joining the thread
/// that blocked in `read(0)`).
#[test]
fn calculate_does_not_wait_for_the_standard_input() {
    use std::time::{Duration, Instant};
    for ext in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("recalculated.{ext}"));
        let mut child = Command::new(SHEETC)
            .args([
                "calculate",
                "-i",
                path(&asset(&format!("formulas.{ext}"))),
                "--output",
                path(&out),
            ])
            .stdin(Stdio::piped())
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .spawn()
            .unwrap();
        // Open, never written and never closed while the command runs.
        let _open = child.stdin.take();
        let started = Instant::now();
        let status = loop {
            if let Some(status) = child.try_wait().unwrap() {
                break status;
            }
            if started.elapsed() > Duration::from_secs(120) {
                let _ = child.kill();
                let _ = child.wait();
                panic!("{ext}: the command waits for the standard input, which it does not take");
            }
            std::thread::sleep(Duration::from_millis(50));
        };
        assert!(status.success(), "{ext}: {status}");
        assert!(out.exists(), "{ext}");
    }
}

#[test]
fn incompatible_output_flags_are_refused_naming_the_conflict_and_the_workbook_comes_from_input_or_file()
 {
    let dir = tempfile::tempdir().unwrap();
    let template = asset("formulas.xlsx");
    let out = dir.path().join("o.xlsx");
    let cases: [(&[&str], &str); 5] = [
        (
            &["--output", path(&out), "--in-place"],
            "both say where the workbook goes",
        ),
        (&["--json"], "--json adds the JSON to a workbook"),
        (
            &["--output", path(&out), "--serve"],
            "does not write a workbook",
        ),
        (&["--output", path(&out), "--pretty"], "add --json"),
        (
            &["--output", path(&template)],
            "never overwrites the original",
        ),
    ];
    for (flags, wanted) in cases {
        let mut args = vec!["calculate", "-i", path(&template)];
        args.extend_from_slice(flags);
        let (code, _, err) = run(SHEETC, &args, "{}");
        assert_eq!(code, 2, "{flags:?}: {err}");
        assert!(err.contains(wanted), "{flags:?}: {err}");
    }
    // The workbook is `-i`, `--input` or `--file`; a positional file is not one.
    for flag in ["--input", "--file"] {
        let (code, stdout, err) = run(SHEETC, &["calculate", flag, path(&template)], "");
        assert_eq!(code, 0, "{flag}: {err}");
        assert!(stdout.contains("valueRanges"), "{flag}");
    }
    let (code, _, err) = run(SHEETC, &["calculate", path(&template)], "");
    assert_eq!(code, 2, "{err}");
    assert!(err.contains("the workbook is missing"), "{err}");
}

#[test]
fn in_place_replaces_the_workbook_and_a_failure_leaves_it_as_it_was() {
    for ext in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let book = dir.path().join(format!("book.{ext}"));
        std::fs::copy(asset(&format!("formulas.{ext}")), &book).unwrap();
        let (code, stdout, err) = run(
            SHEETC,
            &["calculate", "-i", path(&book), "--in-place", "-v"],
            r#"{"inputs": {"T!A1": 10}}"#,
        );
        assert_eq!(code, 0, "{ext}: {err}");
        assert!(stdout.is_empty(), "{ext}: {stdout}");
        assert_eq!(value(&book, "T", 0, 0), CellValue::Number(10.0), "{ext}");
        assert_eq!(value(&book, "F", 0, 0), CellValue::Number(20.0), "{ext}");
        assert!(
            err.contains(&format!("wrote {}", path(&book))),
            "{ext}: {err}"
        );
        // Nothing is left beside it.
        let others: Vec<_> = std::fs::read_dir(dir.path())
            .unwrap()
            .map(|e| e.unwrap().file_name())
            .collect();
        assert_eq!(others.len(), 1, "{ext}: {others:?}");
        // A request that is refused leaves the workbook as it was.
        let before = std::fs::read(&book).unwrap();
        let (code, _, err) = run(
            SHEETC,
            &["calculate", "-i", path(&book), "--in-place"],
            r#"{"inputs": {"F!A1": 10}}"#,
        );
        assert_eq!(code, 1, "{ext}");
        assert!(err.contains("not an input"), "{ext}: {err}");
        assert_eq!(std::fs::read(&book).unwrap(), before, "{ext}");
        assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1, "{ext}");
    }
}

#[test]
fn a_failed_run_does_not_remove_an_output_that_was_there() {
    let dir = tempfile::tempdir().unwrap();
    let template = asset("formulas.xlsx");
    let out = dir.path().join("already.xlsx");
    std::fs::write(&out, b"what the user had").unwrap();
    let (code, _, _) = run(
        SHEETC,
        &["calculate", "-i", path(&template), "--output", path(&out)],
        r#"{"inputs": {"F!A1": 10}}"#,
    );
    assert_eq!(code, 1);
    assert_eq!(std::fs::read(&out).unwrap(), b"what the user had");
    // One that this run created is removed when the run fails, as nothing complete was written.
    let new = dir.path().join("new.xlsx");
    let (code, _, _) = run(
        SHEETC,
        &["calculate", "-i", path(&template), "--output", path(&new)],
        r#"{"inputs": {"F!A1": 10}}"#,
    );
    assert_eq!(code, 1);
    assert!(!new.exists());
}

#[test]
fn an_executable_carries_its_workbook_and_gives_it_json_or_both() {
    for ext in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let exe = dir.path().join("model");
        let template = asset(&format!("formulas.{ext}"));
        let (code, _, err) = run(
            SHEETC,
            &[
                "build",
                "-i",
                path(&template),
                "--embed-template",
                "-o",
                path(&exe),
            ],
            "",
        );
        assert_eq!(code, 0, "{ext}: {err}");

        // `--output` alone: the workbook, and nothing on stdout.
        let out = dir.path().join(format!("a.{ext}"));
        let (code, stdout, err) = run(
            path(&exe),
            &["--output", path(&out)],
            r#"{"inputs": {"T!A1": 5}}"#,
        );
        assert_eq!(code, 0, "{ext}: {err}");
        assert!(stdout.is_empty(), "{ext}: {stdout}");
        assert_eq!(value(&out, "F", 0, 0), CellValue::Number(10.0), "{ext}");

        // No `--output`: the JSON, as any executable gives it.
        let (code, stdout, err) = run(
            path(&exe),
            &["--range", "F!A1"],
            r#"{"inputs": {"T!A1": 5}}"#,
        );
        assert_eq!(code, 0, "{ext}: {err}");
        let json: serde_json::Value = serde_json::from_str(&stdout).unwrap();
        assert_eq!(
            json["valueRanges"][0]["values"],
            serde_json::json!([[10]]),
            "{ext}"
        );
        // Both, from one calculation: the workbook and the same JSON.
        let both = dir.path().join(format!("b.{ext}"));
        let (code, stdout, err) = run(
            path(&exe),
            &["--output", path(&both), "--json", "--range", "F!A1"],
            r#"{"inputs": {"T!A1": 5}}"#,
        );
        assert_eq!(code, 0, "{ext}: {err}");
        let json: serde_json::Value = serde_json::from_str(&stdout).unwrap();
        assert_eq!(
            json["valueRanges"][0]["values"],
            serde_json::json!([[10]]),
            "{ext}"
        );
        assert_eq!(value(&both, "F", 0, 0), CellValue::Number(10.0), "{ext}");
        assert_eq!(value(&both, "F", 3, 0), CellValue::Number(6.0), "{ext}");
        // The JSON of requests, one per line, does not write a workbook.
        let (code, _, err) = run(path(&exe), &["--output", path(&out), "--serve"], "{}");
        assert_eq!(code, 2, "{ext}");
        assert!(err.contains("does not write a workbook"), "{ext}: {err}");
        // The metadata is answered as in any executable, and says what is exposed.
        let (code, meta, _) = run(path(&exe), &["--meta"], "");
        assert_eq!(code, 0, "{ext}");
        let meta: serde_json::Value = serde_json::from_str(&meta).unwrap();
        assert_eq!(meta["mode"], "embedded", "{ext}");
        let rows = meta["meta"].as_array().unwrap();
        assert!(rows.iter().any(|r| r["meta_type"] == "input"), "{ext}");
        assert!(rows.iter().any(|r| r["meta_type"] == "output"), "{ext}");
    }
}

#[test]
fn an_executable_built_with_declarations_answers_them_in_json_and_writes_them_in_the_workbook() {
    let dir = tempfile::tempdir().unwrap();
    let exe = dir.path().join("model");
    let template = asset("formulas.xlsx");
    let (code, _, err) = run(
        SHEETC,
        &[
            "build",
            "-i",
            path(&template),
            "--embed-template",
            "--settings-file",
            &settings(
                dir.path(),
                r#"{"inputs": [{"name": "base", "range": "T!A1"}], "outputs": [{"name": "first", "range": "F!A1"}]}"#,
            ),
            "-o",
            path(&exe),
        ],
        "",
    );
    assert_eq!(code, 0, "{err}");
    let (code, stdout, err) = run(path(&exe), &[], r#"{"inputs": {"base": 4}}"#);
    assert_eq!(code, 0, "{err}");
    let json: serde_json::Value = serde_json::from_str(&stdout).unwrap();
    assert_eq!(
        json,
        serde_json::json!({"valueRanges": [{"name": "first", "range": "F!A1", "values": [[8]]}]})
    );
    let out = dir.path().join("a.xlsx");
    let (code, stdout, err) = run(
        path(&exe),
        &["--output", path(&out), "--json"],
        r#"{"inputs": {"base": 4}}"#,
    );
    assert_eq!(code, 0, "{err}");
    let json: serde_json::Value = serde_json::from_str(&stdout).unwrap();
    assert_eq!(
        json,
        serde_json::json!({"valueRanges": [{"name": "first", "range": "F!A1", "values": [[8]]}]})
    );
    assert_eq!(value(&out, "T", 0, 0), CellValue::Number(4.0));
    assert_eq!(value(&out, "F", 0, 0), CellValue::Number(8.0));
    // `F!A4` is not an output: it keeps the result the template stored.
    assert_eq!(value(&out, "F", 3, 0), value(&template, "F", 3, 0));
    // The declared inputs and outputs are what `--meta` lists.
    let (_, meta, _) = run(path(&exe), &["--meta"], "");
    assert!(meta.contains("\"base\""), "{meta}");
    assert!(meta.contains("\"first\""), "{meta}");
}

#[test]
fn an_executable_with_nothing_given_calculates_every_formula_again() {
    let dir = tempfile::tempdir().unwrap();
    let exe = dir.path().join("model");
    let template = asset("formulas.xlsx");
    let (code, _, err) = run(
        SHEETC,
        &[
            "build",
            "-i",
            path(&template),
            "--embed-template",
            "-o",
            path(&exe),
        ],
        "",
    );
    assert_eq!(code, 0, "{err}");
    let out = dir.path().join("a.xlsx");
    let (code, _, err) = run(path(&exe), &["--output", path(&out)], "");
    assert_eq!(code, 0, "{err}");
    assert_eq!(value(&out, "F", 0, 0), CellValue::Number(2.0));
    // Inputs can be given: the same executable answers with them.
    let (code, _, err) = run(
        path(&exe),
        &["--output", path(&out)],
        r#"{"inputs": {"T!A1": 5}}"#,
    );
    assert_eq!(code, 0, "{err}");
    assert_eq!(value(&out, "F", 0, 0), CellValue::Number(10.0));
    // It has no file to replace.
    let (code, _, err) = run(path(&exe), &["--in-place"], "");
    assert_eq!(code, 2);
    assert!(err.contains("--in-place"), "{err}");
    let (_, meta, _) = run(path(&exe), &["--meta"], "");
    assert!(meta.contains("\"embedded\""), "{meta}");
    assert!(!meta.contains("recalculate"), "{meta}");
}

#[test]
fn an_executable_without_its_workbook_does_not_write_one() {
    let dir = tempfile::tempdir().unwrap();
    let exe = dir.path().join("json");
    let template = asset("formulas.xlsx");
    let settings = dir.path().join("settings.json");
    std::fs::write(
        &settings,
        r#"{"inputs": [{"name": "base", "range": "T!A1"}], "outputs": [{"name": "first", "range": "F!A1"}]}"#,
    )
    .unwrap();
    let (code, _, err) = run(
        SHEETC,
        &[
            "build",
            "-i",
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

/// `=SQRT(-1)` is `Err:502` in `LibreOffice`, and the circular reference of `cells.*` is `Err:522`. A
/// recalculated workbook holds them as the format can: an ODS keeps the code, an XLSX has none and
/// holds `#VALUE!`, which is what `LibreOffice`'s own export writes.
#[test]
fn a_workbook_calculated_again_holds_the_errors_as_the_format_can() {
    use sheetrs_io::reader::ExcelError;
    for (ext, sqrt, circular) in [
        (
            "xlsx",
            CellValue::Error(ExcelError::Value),
            CellValue::Error(ExcelError::Value),
        ),
        (
            "ods",
            CellValue::Error(ExcelError::Libre(502)),
            CellValue::Error(ExcelError::Libre(522)),
        ),
    ] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("recalculated.{ext}"));
        let (code, _, err) = run(
            SHEETC,
            &[
                "calculate",
                "-i",
                path(&asset(&format!("cells.{ext}"))),
                "--output",
                path(&out),
            ],
            "",
        );
        assert_eq!(code, 0, "{ext}: {err}");
        assert_eq!(value(&out, "T", 13, 0), sqrt, "{ext}: SQRT(-1)");
        assert_eq!(value(&out, "T", 15, 0), circular, "{ext}: circular");
    }
}

/// What the command wrote is information, not a warning: it is on the standard error with `-v`, and
/// with `--log-format json` it is a record with a stable code and the path and the cells as context.
#[test]
fn what_the_command_wrote_is_information_given_with_the_verbose_flag_and_as_a_json_record() {
    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join("out.xlsx");
    let book = asset("formulas.xlsx");
    let request = r#"{"inputs": {"T!A1": 10}}"#;
    let base = ["calculate", "-i", path(&book), "-o", path(&out)];
    let with = |extra: &[&'static str]| [&base[..], extra].concat();
    let (code, _, quiet) = run(SHEETC, &with(&[]), request);
    assert_eq!(code, 0, "{quiet}");
    assert!(
        !quiet.contains("wrote "),
        "what was written is said with -v: {quiet}"
    );
    let (_, _, verbose) = run(SHEETC, &with(&["-v"]), request);
    assert!(
        verbose
            .lines()
            .any(|line| line.starts_with(&format!("wrote {}", path(&out)))),
        "{verbose}"
    );
    let (_, _, json) = run(SHEETC, &with(&["-v", "--log-format", "json"]), request);
    // The records of the run, one per line: the one of what was written is the information one.
    let record = json
        .lines()
        .map(|line| serde_json::from_str::<serde_json::Value>(line).unwrap())
        .find(|record| record["code"] == "WORKBOOK_WRITTEN")
        .unwrap_or_else(|| panic!("no record of what was written: {json}"));
    assert_eq!(record["level"], "info");
    assert_eq!(record["context"]["path"], path(&out));
    let cells = record["context"]["cells"].as_str().unwrap();
    assert!(cells.parse::<usize>().unwrap() > 0);
}

//! The tools of `sheetc` in memory (`sheetrs_compiler::memory`) give what the command line gives:
//! the same answers from bytes and JSON as from files and flags.

use serde_json::{Value as Json, json};
use sheetrs_compiler::memory::BookCompile;
use sheetrs_compiler::memory::{self, Options};
use sheetrs_engine::ir::Mode;
use sheetrs_io::Book;
use std::path::{Path, PathBuf};
use std::process::Command;

const SHEETC: &str = env!("CARGO_BIN_EXE_sheetc");

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(name)
}

fn bytes_of(name: &str) -> Book {
    Book::from_bytes(std::fs::read(asset(name)).unwrap()).unwrap()
}

/// `sheetc` with `args`, the JSON `request` on its standard input; the standard output.
fn sheetc(args: &[&str], request: &str) -> String {
    use std::io::Write;
    // The request is read from the standard input only when it is asked for: `--request -`.
    let mut args = args.to_vec();
    if !request.is_empty() && !args.contains(&"--serve") && !args.contains(&"--request") {
        args.extend(["--request", "-"]);
    }
    let mut child = Command::new(SHEETC)
        .args(&args)
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .stderr(std::process::Stdio::piped())
        .spawn()
        .unwrap();
    child
        .stdin
        .take()
        .unwrap()
        .write_all(request.as_bytes())
        .unwrap();
    let out = child.wait_with_output().unwrap();
    assert!(
        out.status.success(),
        "{}",
        String::from_utf8_lossy(&out.stderr)
    );
    String::from_utf8(out.stdout).unwrap()
}

const REQUEST: &str =
    r#"{"inputs": {"principal": 250000, "rate": 0.045, "years": 15, "month": 24}}"#;

#[test]
fn values_in_memory_are_the_values_of_the_command() {
    for file in ["loan.xlsx", "loan.ods"] {
        let path = asset(file);
        // Compared as text: parsing a number back is not exact in every build of serde_json.
        let cli = sheetc(&["calculate", "-i", path.to_str().unwrap()], REQUEST);
        let request: Json = serde_json::from_str(REQUEST).unwrap();
        let done = bytes_of(file)
            .calculate(&request, &Options::default(), Mode::Values, false)
            .unwrap();
        assert_eq!(
            done.answer.unwrap().to_json().to_string(),
            cli.trim(),
            "{file}"
        );
        assert!(done.workbook.is_none());
    }
}

#[test]
fn the_workbook_in_memory_is_the_workbook_of_the_command() {
    let dir = tempfile::tempdir().unwrap();
    for (file, request) in [
        ("loan.xlsx", REQUEST),
        ("loan.ods", REQUEST),
        // Nothing given: every formula calculated again.
        ("loan.xlsx", "{}"),
    ] {
        let path = asset(file);
        let out = dir.path().join(format!("{}-{file}", request.len()));
        sheetc(
            &[
                "calculate",
                "-i",
                path.to_str().unwrap(),
                "--output",
                out.to_str().unwrap(),
            ],
            request,
        );
        let done = bytes_of(file)
            .calculate(
                &serde_json::from_str(request).unwrap(),
                &Options::default(),
                Mode::Embedded,
                false,
            )
            .unwrap();
        assert_eq!(
            done.workbook.unwrap(),
            std::fs::read(&out).unwrap(),
            "{file} {request}"
        );
        assert!(done.written > 0);
    }
}

#[test]
fn the_workbook_and_the_json_come_from_one_calculation() {
    let request: Json = serde_json::from_str(REQUEST).unwrap();
    let both = bytes_of("loan.xlsx")
        .calculate(&request, &Options::default(), Mode::Embedded, true)
        .unwrap();
    let alone = bytes_of("loan.xlsx")
        .calculate(&request, &Options::default(), Mode::Embedded, false)
        .unwrap();
    assert_eq!(both.workbook, alone.workbook);
    let answer = both.answer.unwrap().to_json();
    assert!(
        answer["valueRanges"][0]["values"][0][0].is_number(),
        "{answer}"
    );
    assert_eq!(answer["valueRanges"][0]["name"], "payment", "{answer}");
}

#[test]
fn a_recalculated_workbook_gives_no_json() {
    let error = bytes_of("loan.xlsx")
        .calculate(&json!({}), &Options::default(), Mode::Recalculate, true)
        .unwrap_err();
    assert!(error.to_string().contains("no JSON"), "{error}");
}

#[test]
fn inspect_in_memory_says_what_the_command_says() {
    let report = bytes_of("loan.xlsx").inspect(&Options::default()).unwrap();
    let text = sheetc(&["inspect", "-i", asset("loan.xlsx").to_str().unwrap()], "");
    assert!(text.contains(&format!("sheets: {}", report.sheets.join(", "))));
    assert!(text.contains(&format!(
        "cells: {} ({} formulas)",
        report.cells, report.formulas
    )));
    assert!(report.unsupported_functions.is_empty());
    assert!(report.functions.iter().all(|f| f.implemented));
}

#[test]
fn evaluate_in_memory_is_the_value_the_command_prints() {
    for formula in ["=SUM(1,2,3)", "=TEXT(0.25,\"0%\")", "=1/0"] {
        let printed = sheetc(&["evaluate", formula], "");
        let value = memory::evaluate(formula, None, &Options::default()).unwrap();
        assert_eq!(value.to_string(), printed.trim(), "{formula}");
    }
}

#[test]
fn verify_in_memory_matches_every_cached_formula_of_the_loan() {
    let result = bytes_of("loan.xlsx")
        .verify(&Options::default(), &[], None, None)
        .unwrap();
    assert_eq!(result["mismatches"], 0, "{result}");
    assert!(result["sheets"][0]["formulas"].as_u64().unwrap() > 0);
}

#[test]
fn the_options_are_the_keys_of_a_configuration_file() {
    let options = Options::from_json(&json!({
        "now": 45000.5,
        "locale": {"tag": "es-CO"},
        "declarations": {"outputs": [{"name": "payment", "range": "Loan!B6"}]},
    }))
    .unwrap();
    assert_eq!(options.config.now, Some(45000.5));
    assert!(options.declarations.is_some());
    // A file cannot be read here, and an unknown key is a mistake.
    assert!(Options::from_json(&json!({"settings_file": "model.json"})).is_err());
    assert!(Options::from_json(&json!({"bogus": 1})).is_err());
    // Folding is a Rust option of the compiler, not a key of the configuration.
    assert!(Options::from_json(&json!({"fold": false})).is_err());
    assert!(Options::from_json(&Json::Null).is_ok());
}

#[test]
fn declarations_given_as_json_replace_the_settings_sheet() {
    let options = Options::from_json(&json!({
        "declarations": {
            "inputs": [{"name": "rate", "range": "Loan!B2"}],
            "outputs": [{"name": "payment", "range": "Loan!B6"}]
        }
    }))
    .unwrap();
    let done = bytes_of("loan.xlsx")
        .calculate(
            &json!({"inputs": {"rate": 0.05}}),
            &options,
            Mode::Values,
            false,
        )
        .unwrap();
    let answer = done.answer.unwrap().to_json();
    assert!(
        answer["valueRanges"][0]["values"][0][0].is_number(),
        "{answer}"
    );
    assert_eq!(answer["valueRanges"][0]["name"], "payment", "{answer}");
}

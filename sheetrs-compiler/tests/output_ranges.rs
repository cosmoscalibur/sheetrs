//! Which `ranges` of a request a calculation answers, in an XLSX and in an ODS (`loan.*` declares the
//! outputs `payment` = `Loan!B4`, `interest` = `Loan!B6` and `balance` = `Loan!B8`, and the inputs
//! `principal` = `Loan!B1`, ...; `Loan!B5` is a formula that nothing declares and `Loan!B2` a constant).
//!
//! - The compiled program (`Mode::Values`): the declared outputs are all it answers; `ranges` selects
//!   among them and a range outside is an error that lists them. With none declared the outputs are the
//!   formula cells.
//! - The embedded mode (the workbook is its own template): a range outside the outputs is answered
//!   from the workbook as it was written.

use serde_json::{Value as Json, json};
use sheetrs_compiler::memory::{BookCompile, Options};
use sheetrs_compiler::{Error, Result};
use sheetrs_engine::ir::Mode;
use sheetrs_io::Book;
use std::path::{Path, PathBuf};

const FILES: [&str; 2] = ["loan.xlsx", "loan.ods"];

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(name)
}

fn loan(name: &str) -> Book {
    Book::from_bytes(std::fs::read(asset(name)).unwrap()).unwrap()
}

/// The loan with no output declared (and no input): its outputs are its formula cells.
fn undeclared() -> Options {
    Options::from_json(&json!({"declarations": {}})).unwrap()
}

fn values(answer: &Json) -> Vec<(String, Json)> {
    answer["valueRanges"]
        .as_array()
        .unwrap()
        .iter()
        .map(|r| {
            (
                r["range"].as_str().unwrap().to_string(),
                r["values"].clone(),
            )
        })
        .collect()
}

fn answer(book: &Book, request: &Json, options: &Options, mode: Mode) -> Result<Json> {
    let calculated = book.calculate(request, options, mode, mode.is_workbook())?;
    Ok(serde_json::to_value(calculated.answer.unwrap()).unwrap())
}

#[test]
fn declared_and_inside_the_outputs_is_a_selection() {
    for file in FILES {
        let asked = json!({"ranges": ["Loan!B6", "payment"]});
        let answer = answer(&loan(file), &asked, &Options::default(), Mode::Values).unwrap();
        let got = values(&answer);
        assert_eq!(got.len(), 2, "{file}: {answer}");
        assert_eq!(got[0].0, "Loan!B6", "{file}");
        assert_eq!(got[1].0, "Loan!B4", "{file}");
        assert!((got[1].1[0][0].as_f64().unwrap() - 1_073.643_246_024_279_5).abs() < 1e-6);
    }
}

#[test]
fn declared_and_outside_the_outputs_is_an_error_that_lists_the_declared_ones() {
    for file in FILES {
        for outside in ["Loan!B5", "Loan!B2", "Loan!B4:B5"] {
            let asked = json!({"ranges": [outside]});
            let error = answer(&loan(file), &asked, &Options::default(), Mode::Values).unwrap_err();
            let Error::Engine(sheetrs_engine::Error::OutsideOutputs { range, declared }) = &error
            else {
                panic!("{file} {outside}: not the typed error: {error:?}");
            };
            assert_eq!(range, outside, "{file}");
            assert_eq!(
                declared.as_deref(),
                Some(
                    &[
                        "payment (Loan!B4)".to_string(),
                        "interest (Loan!B6)".to_string(),
                        "balance (Loan!B8)".to_string(),
                    ][..]
                ),
                "{file}"
            );
            let said = error.to_string();
            assert!(said.contains(outside), "{said}");
            assert!(said.contains("payment (Loan!B4)"), "{said}");
            assert!(said.contains("balance (Loan!B8)"), "{said}");
        }
    }
}

#[test]
fn with_no_output_declared_the_ranges_are_among_the_formula_cells() {
    for file in FILES {
        let book = loan(file);
        let inside = answer(
            &book,
            &json!({"ranges": ["Loan!B5"]}),
            &undeclared(),
            Mode::Values,
        )
        .unwrap();
        // 1073.64 a month, 360 months.
        assert!((values(&inside)[0].1[0][0].as_f64().unwrap() - 386_511.5).abs() < 1.0);
        // A constant is not an output, and with none declared the error has no list to give.
        let error = answer(
            &book,
            &json!({"ranges": ["Loan!B2"]}),
            &undeclared(),
            Mode::Values,
        )
        .unwrap_err();
        let Error::Engine(sheetrs_engine::Error::OutsideOutputs { declared, .. }) = &error else {
            panic!("{file}: {error:?}");
        };
        assert!(declared.is_none(), "{file}");
        assert!(error.to_string().contains("formula cells"), "{error}");
    }
}

#[test]
fn embedded_and_declared_a_range_outside_the_outputs_is_answered_from_the_written_workbook() {
    for file in FILES {
        // A formula nothing declares (`Total paid`, with the principal of the request), a constant
        // (`Annual rate`), and a declared output between them: the answer keeps the order asked.
        let asked = json!({
            "inputs": {"principal": 1000},
            "ranges": ["Loan!B5", "payment", "Loan!B2"],
        });
        let answer = answer(&loan(file), &asked, &Options::default(), Mode::Embedded).unwrap();
        let got = values(&answer);
        let payment = got[1].1[0][0].as_f64().unwrap();
        assert_eq!(
            got.iter().map(|g| g.0.as_str()).collect::<Vec<_>>(),
            ["Loan!B5", "Loan!B4", "Loan!B2"],
            "{file}: {answer}"
        );
        assert!(
            (payment - 5.368_216_230_121_398).abs() < 1e-9,
            "{file}: {payment}"
        );
        // Total paid = payment * 30 * 12 with the new principal: calculated, not the cached value.
        let total = got[0].1[0][0].as_f64().unwrap();
        assert!((total - payment * 360.0).abs() < 1e-6, "{file}: {total}");
        assert!(
            (got[2].1[0][0].as_f64().unwrap() - 0.05).abs() < 1e-12,
            "{file}"
        );
    }
}

#[test]
fn embedded_a_range_that_is_not_one_of_the_workbook_is_still_an_error() {
    for file in FILES {
        let asked = json!({"ranges": ["Nowhere!A1"]});
        let error = answer(&loan(file), &asked, &Options::default(), Mode::Embedded).unwrap_err();
        assert!(error.to_string().contains("Nowhere!A1"), "{file}: {error}");
    }
}

#[test]
fn embedded_with_none_declared_a_range_of_constants_is_read_from_the_workbook() {
    for file in FILES {
        let asked = json!({"ranges": ["Loan!A1:B3"]});
        let answer = answer(&loan(file), &asked, &undeclared(), Mode::Embedded).unwrap();
        let got = values(&answer);
        assert_eq!(got[0].0, "Loan!A1:B3", "{file}");
        assert_eq!(got[0].1[0][0], json!("Principal"), "{file}: {answer}");
        assert_eq!(got[0].1[0][1].as_f64(), Some(200_000.0), "{file}");
        assert_eq!(got[0].1[2][1].as_f64(), Some(30.0), "{file}");
    }
}

#[test]
fn embedded_the_answer_of_the_declared_outputs_is_what_it_was() {
    for file in FILES {
        let answer = answer(&loan(file), &json!({}), &Options::default(), Mode::Embedded).unwrap();
        let names: Vec<_> = answer["valueRanges"]
            .as_array()
            .unwrap()
            .iter()
            .map(|r| r["name"].as_str().unwrap().to_string())
            .collect();
        assert_eq!(names, ["payment", "interest", "balance"], "{file}");
    }
}

/// `sheetc calculate -i FILE` with `extra`: the exit code, the standard output and the standard error.
fn sheetc(file: &str, extra: &[&str]) -> (i32, String, String) {
    let out = std::process::Command::new(env!("CARGO_BIN_EXE_sheetc"))
        .args(["calculate", "-i"])
        .arg(asset(file))
        .args(extra)
        .output()
        .unwrap();
    (
        out.status.code().unwrap_or(-1),
        String::from_utf8_lossy(&out.stdout).into_owned(),
        String::from_utf8_lossy(&out.stderr).into_owned(),
    )
}

#[test]
fn the_command_refuses_a_range_outside_the_declared_outputs_and_lists_them() {
    for file in FILES {
        let (code, _, err) = sheetc(file, &["--range", "Loan!B5"]);
        assert_eq!(code, 1, "{file}: {err}");
        assert!(err.contains("`Loan!B5` is outside the outputs"), "{err}");
        assert!(err.contains("payment (Loan!B4)"), "{err}");
    }
}

#[test]
fn the_command_that_writes_the_workbook_answers_a_range_outside_the_declared_outputs() {
    let dir = tempfile::tempdir().unwrap();
    for file in FILES {
        let out = dir.path().join(format!("out-{file}"));
        let (code, answer, err) = sheetc(
            file,
            &["--json", "--range", "Loan!B5", "-o", out.to_str().unwrap()],
        );
        assert_eq!(code, 0, "{file}: {err}");
        let answer: Json = serde_json::from_str(&answer).unwrap();
        assert_eq!(values(&answer)[0].0, "Loan!B5", "{file}");
        assert!((values(&answer)[0].1[0][0].as_f64().unwrap() - 386_511.5).abs() < 1.0);
        assert!(out.is_file(), "{file}");
    }
}

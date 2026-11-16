//! The `ranges` of a `calculate` step, in an XLSX and in an ODS. The workbook is its own template here (the
//! embedded mode): the outputs it declares are a selection for `ranges`, and a range outside them is read, in
//! memory and with no recalculation, from the workbook just written, as `query` reads any workbook (the
//! compiled program, which writes none, refuses it: see `sheetrs-compiler/tests/output_ranges.rs`).
//!
//! `loan.*` declares `payment` = `Loan!B4`, `interest` = `Loan!B6` and `balance` = `Loan!B8` as outputs;
//! `Loan!B5` is a formula nothing declares and `Loan!B2` a constant. `structure/book.*` declares nothing.

use serde_json::{Value as Json, json};
use sheetrs_edit::{BookEdit, Calculation, Error};
use sheetrs_io::Book;
use sheetrs_query::{BookQuery, RangeOptions};
use std::path::{Path, PathBuf};

const EXTENSIONS: [&str; 2] = ["xlsx", "ods"];

fn loan(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-compiler/tests/assets")
        .join(format!("loan.{extension}"))
}

fn structure(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/structure")
        .join(format!("book.{extension}"))
}

/// The answer of a calculation of `file` with `inputs` and `ranges`: `{range: values}` in the order asked,
/// or the error of the run.
fn ask(file: &Path, inputs: &Json, ranges: &[&str]) -> Result<Vec<(String, Json)>, Error> {
    let mut book = Book::from_file(file).unwrap();
    book.calculate(Calculation {
        inputs: inputs.as_object().cloned().unwrap_or_default(),
        ranges: Some(ranges.iter().map(|r| (*r).to_string()).collect()),
        answer: true,
        ..Calculation::default()
    });
    let report = book.write_to(std::io::sink())?;
    let answer = report.value_ranges.unwrap().to_json();
    Ok(answer["valueRanges"]
        .as_array()
        .unwrap()
        .iter()
        .map(|r| {
            (
                r["range"].as_str().unwrap().to_string(),
                r["values"].clone(),
            )
        })
        .collect())
}

fn number(values: &Json, row: usize, column: usize) -> f64 {
    values[row][column].as_f64().unwrap()
}

#[test]
fn declared_and_inside_the_outputs_is_a_selection() {
    for extension in EXTENSIONS {
        let got = ask(&loan(extension), &json!({}), &["Loan!B6", "payment"]).unwrap();
        assert_eq!(got.len(), 2, "{extension}");
        assert_eq!(got[0].0, "Loan!B6");
        assert_eq!(got[1].0, "Loan!B4");
        assert!((number(&got[1].1, 0, 0) - 1_073.643_246_024_279_5).abs() < 1e-6);
    }
}

#[test]
fn declared_and_outside_the_outputs_is_read_from_the_written_workbook() {
    for extension in EXTENSIONS {
        // The formula nothing declares keeps the result the template stored; the constant and the input
        // given are in the workbook as written.
        let got = ask(
            &loan(extension),
            &json!({"principal": 1000}),
            &["Loan!B5", "payment", "Loan!B2", "Loan!B1"],
        )
        .unwrap();
        let ranges: Vec<_> = got.iter().map(|g| g.0.as_str()).collect();
        assert_eq!(
            ranges,
            ["Loan!B5", "Loan!B4", "Loan!B2", "Loan!B1"],
            "{extension}"
        );
        assert!(
            (number(&got[1].1, 0, 0) - 5.368_216_230_121_398).abs() < 1e-9,
            "{extension}"
        );
        assert!(
            (number(&got[0].1, 0, 0) - 386_511.5).abs() < 1.0,
            "{extension}"
        );
        assert!(
            (number(&got[2].1, 0, 0) - 0.05).abs() < 1e-12,
            "{extension}"
        );
        assert!(
            (number(&got[3].1, 0, 0) - 1000.0).abs() < 1e-12,
            "{extension}"
        );
    }
}

#[test]
fn declared_a_range_that_is_not_one_of_the_workbook_is_an_error() {
    for extension in EXTENSIONS {
        let error = ask(&loan(extension), &json!({}), &["Nowhere!A1"]).unwrap_err();
        assert!(format!("{error:#}").contains("Nowhere!A1"), "{error:#}");
    }
}

#[test]
fn none_declared_the_ranges_are_the_formula_cells_and_the_constants_the_workbook_holds() {
    for extension in EXTENSIONS {
        // F!A5:A6 are formula cells (the outputs); T!A1:A3 constants, read from the workbook as written.
        let got = ask(&structure(extension), &json!({}), &["F!A5:A6", "T!A1:A3"]).unwrap();
        assert_eq!(got[0].1, json!([[5], [9]]), "{extension}");
        assert_eq!(got[1].0, "T!A1:A3", "{extension}");
        assert_eq!(got[1].1[0][0].as_f64(), Some(1.0), "{extension}: {got:?}");
    }
}

#[test]
fn a_range_outside_the_outputs_is_read_from_the_workbook_written_in_memory() {
    for extension in EXTENSIONS {
        // The calculation writes the workbook (inputs in it, declared outputs calculated); what lies
        // outside the outputs is read from those bytes with `query`, with no recalculation.
        let mut book = Book::from_file(loan(extension)).unwrap();
        book.calculate(Calculation {
            inputs: json!({"principal": 1000}).as_object().cloned().unwrap(),
            ..Calculation::default()
        });
        let mut bytes = Vec::new();
        book.write_to(&mut bytes).unwrap();
        let written = Book::from_bytes(bytes).unwrap();
        let blocks = written
            .range(["Loan!B1", "Loan!B2", "Loan!B4"], &RangeOptions::default())
            .unwrap();
        let value = |block: usize| {
            blocks[block].values.as_ref().unwrap()[0][0]
                .as_ref()
                .and_then(Json::as_f64)
                .unwrap()
        };
        assert!((value(0) - 1000.0).abs() < 1e-12, "{extension}: the input");
        assert!((value(1) - 0.05).abs() < 1e-12, "{extension}: a constant");
        assert!(
            (value(2) - 5.368_216_230_121_398).abs() < 1e-9,
            "{extension}: an output, calculated"
        );
    }
}

#[test]
fn an_input_is_in_the_workbook_a_range_outside_the_outputs_reads() {
    for extension in EXTENSIONS {
        let mut book = Book::from_file(structure(extension)).unwrap();
        book.calculate(Calculation {
            inputs: json!({"T!A1": 10}).as_object().cloned().unwrap(),
            ..Calculation::default()
        });
        let mut bytes = Vec::new();
        book.write_to(&mut bytes).unwrap();
        let blocks = Book::from_bytes(bytes)
            .unwrap()
            .range(["T!A1"], &RangeOptions::default())
            .unwrap();
        let first = blocks[0].values.as_ref().unwrap()[0][0].as_ref();
        assert_eq!(first.and_then(Json::as_f64), Some(10.0), "{extension}");
    }
}

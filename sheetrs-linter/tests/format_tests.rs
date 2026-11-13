//! The three formats of a run on a real workbook, through the library (the executable is tested in
//! `sheetrs-cli`): they must say the same findings.

use sheetrs_io::Book;
use sheetrs_linter::output::render_report;
use sheetrs_linter::{Linter, LinterConfig, OutputFormat};

const XLSX: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/lint_book.xlsx");

fn run(format: OutputFormat) -> (String, usize) {
    let workbook = Book::from_file(XLSX).unwrap().workbook().unwrap();
    let config = LinterConfig::default();
    let violations = Linter::with_config(config.clone()).lint_workbook(&workbook);
    let text = render_report(format, XLSX, &violations, &workbook, &config, false).unwrap();
    (text, violations.len())
}

#[test]
fn json_summary_and_violations_agree_with_the_run() {
    let (text, count) = run(OutputFormat::Json);
    let doc: serde_json::Value = serde_json::from_str(&text).unwrap();
    assert_eq!(doc["file"], XLSX);
    assert_eq!(doc["summary"]["total"], count);
    assert_eq!(doc["violations"].as_array().unwrap().len(), count);
    let sum = ["errors", "warnings", "info"]
        .iter()
        .map(|k| doc["summary"][k].as_u64().unwrap())
        .sum::<u64>();
    assert_eq!(sum, count as u64);
    let first = &doc["violations"][0];
    for key in ["message", "rule_id", "scope", "severity"] {
        assert!(first.get(key).is_some(), "{key}");
    }
}

#[test]
fn the_three_formats_report_the_same_number_of_findings() {
    let (json, count) = run(OutputFormat::Json);
    let (sarif, _) = run(OutputFormat::Sarif);
    let (human, _) = run(OutputFormat::Human);
    let sarif: serde_json::Value = serde_json::from_str(&sarif).unwrap();
    assert_eq!(sarif["runs"][0]["results"].as_array().unwrap().len(), count);
    let json: serde_json::Value = serde_json::from_str(&json).unwrap();
    assert_eq!(json["violations"].as_array().unwrap().len(), count);
    assert!(human.starts_with("Linting: "));
    assert!(human.contains("Summary:"));
    assert!(!human.contains('\u{1b}'), "no styles when color is off");
}

//! The SARIF log of a real workbook against the official SARIF 2.1.0 JSON schema
//! (`assets/sarif-schema-2.1.0.json`, from the OASIS `sarif-spec` repository).

use sheetrs_io::Book;
use sheetrs_linter::sarif::write_sarif;
use sheetrs_linter::{Linter, LinterConfig};

const SCHEMA: &str = include_str!("assets/sarif-schema-2.1.0.json");
const XLSX: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/lint_book.xlsx");
const ODS: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/lint_book.ods");

fn log_of(path: &str) -> (serde_json::Value, usize) {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let config = LinterConfig::default();
    let violations = Linter::with_config(config.clone()).lint_workbook(&workbook);
    let mut out = Vec::new();
    write_sarif(&mut out, "tests/lint_book", &violations, &workbook, &config).unwrap();
    (serde_json::from_slice(&out).unwrap(), violations.len())
}

fn validator() -> jsonschema::Validator {
    let schema: serde_json::Value = serde_json::from_str(SCHEMA).unwrap();
    jsonschema::draft4::new(&schema).unwrap()
}

#[test]
fn the_log_of_each_workbook_is_valid_sarif() {
    let validator = validator();
    for path in [XLSX, ODS] {
        let (log, count) = log_of(path);
        assert!(count > 10, "the workbook is meant to trigger many rules");
        let errors: Vec<String> = validator
            .iter_errors(&log)
            .map(|e| format!("{} at {}", e, e.instance_path()))
            .collect();
        assert!(errors.is_empty(), "{path}: {errors:#?}");
        assert_eq!(log["runs"][0]["results"].as_array().unwrap().len(), count);
    }
}

#[test]
fn every_result_points_to_its_rule_and_a_region() {
    let (log, _) = log_of(XLSX);
    let rules = log["runs"][0]["tool"]["driver"]["rules"]
        .as_array()
        .unwrap();
    for result in log["runs"][0]["results"].as_array().unwrap() {
        let index = usize::try_from(result["ruleIndex"].as_u64().unwrap()).unwrap();
        assert_eq!(rules[index]["id"], result["ruleId"]);
        let region = &result["locations"][0]["physicalLocation"]["region"];
        assert!(region["startLine"].as_u64().unwrap() >= 1);
        assert!(region["startColumn"].as_u64().unwrap() >= 1);
        assert!(matches!(
            result["level"].as_str().unwrap(),
            "error" | "warning" | "note"
        ));
    }
}

/// The schema is not vacuous: a log with a level it does not know is refused.
#[test]
fn the_validator_refuses_what_is_not_sarif() {
    let (mut log, _) = log_of(XLSX);
    log["runs"][0]["results"][0]["level"] = serde_json::json!("fatal");
    assert!(!validator().is_valid(&log));
    log["runs"][0]["results"][0]["level"] = serde_json::json!("error");
    log["version"] = serde_json::json!("2.0.0");
    assert!(!validator().is_valid(&log));
}

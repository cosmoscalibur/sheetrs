//! Workbooks without a settings sheet: the same declarations in a JSON file named by
//! `sheetc.toml`, and the convention to recalculate a workbook as it is.
//!
//! The workbooks were made by `LibreOffice` in both formats (`sheetrs-io/tests/assets`): sheet `T`
//! holds numbers and sheet `F` formulas such as `T!A1*2`.

use serde_json::{Value as Json, json};
use sheetrs_compiler::config::{self, Config};
use sheetrs_compiler::{compile, loader};
use sheetrs_engine::Overrides;
use std::path::{Path, PathBuf};

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-io/tests/assets")
        .join(name)
}

fn model(name: &str) -> sheetrs_compiler::Model {
    loader::load(&sheetrs_io::reader::Source::from_file(asset(name)).unwrap()).unwrap()
}

const SETTINGS: &str = r#"{
  "inputs":  { "base": "T!A1" },
  "outputs": { "doubled": "F!A1:A2" },
  "meta":    [ { "key": "version", "value": "1", "description": "model", "meta_type": "info" } ]
}"#;

fn from_toml(dir: &Path, toml: &str) -> Config {
    std::fs::write(dir.join(config::FILE_NAME), toml).unwrap();
    config::load(&dir.join(config::FILE_NAME)).unwrap()
}

#[test]
fn a_workbook_without_settings_sheet_takes_its_declarations_from_the_json_file() {
    let dir = tempfile::tempdir().unwrap();
    std::fs::write(dir.path().join("model.json"), SETTINGS).unwrap();
    // The path is relative to the configuration file, not to the working directory.
    let config = from_toml(dir.path(), "settings_file = \"model.json\"\n");
    for file in ["formulas.xlsx", "formulas.ods"] {
        let model = model(file);
        let spec = config.spec(&model).unwrap();
        assert_eq!(spec.inputs.len(), 1, "{file}");
        assert_eq!(spec.skip_sheet, None, "no settings sheet to leave out");
        let (program, _) = compile(&model, &spec).unwrap();
        // The input changes: `T!A1` is the base of `F!A1` (`=T!A1*2`).
        let out = program
            .execute(&json!({"base": 10}), &Overrides::default())
            .unwrap();
        assert_eq!(out, json!({"doubled": [[20], [20]]}), "{file}");
        let keys: Vec<_> = program.meta.iter().map(|m| m.key.as_str()).collect();
        assert!(
            keys.ends_with(&["base", "doubled", "version"]),
            "{file}: {keys:?}"
        );
    }
}

#[test]
fn the_json_file_is_checked_as_the_sheet_is() {
    let model = model("formulas.xlsx");
    let spec = |json: &str| {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("m.json");
        std::fs::write(&path, json).unwrap();
        Config {
            settings_file: Some(path),
            ..Config::default()
        }
        .spec(&model)
    };
    let message = |json: &str| spec(json).unwrap_err().to_string();
    assert!(message(r#"{"inputs": {"a": "Nowhere!A1"}}"#).contains("input `a`"));
    assert!(
        message(r#"{"inputs": {"a": "T!A1"}, "outputs": {"A": "T!A2"}}"#)
            .contains("already defined")
    );
    assert!(message(r#"{"inputs": {"T": "T!A1"}}"#).contains("also a sheet"));
    assert!(message(r#"{"inputs": {"a": "A1"}}"#).contains("must start with the sheet"));
    assert!(message(r#"{"inputz": {}}"#).contains("unknown field"));
    assert!(
        message(r#"{"meta": [{"key": "x", "meta_type": "input", "value": "T!A1"}]}"#)
            .contains("under `inputs`")
    );
    assert!(
        spec(
            r#"{"inputs": {"a": {"range": "T!A1", "description": "d"}}, "outputs": {"o": "F!A1"}}"#
        )
        .is_ok()
    );
}

#[test]
fn a_sheet_and_a_file_together_are_refused() {
    let config = Config {
        settings: Some("settings".into()),
        settings_file: Some("m.json".into()),
        ..Config::default()
    };
    let e = config
        .spec(&model("formulas.xlsx"))
        .unwrap_err()
        .to_string();
    assert!(e.contains("name one"), "{e}");
}

/// `outputs = any` with no input is a program like any other: nothing is an input, any range can be
/// asked for, and the answer is what the formulas give. (To recalculate a workbook and get the
/// workbook back, see `template.rs`.)
#[test]
fn outputs_any_without_inputs_answers_any_range() {
    let config = Config {
        outputs: Some("any".into()),
        ..Config::default()
    };
    for file in ["formulas.xlsx", "formulas.ods"] {
        let model = model(file);
        let spec = config.spec(&model).unwrap();
        assert!(spec.inputs.is_empty() && spec.outputs.is_empty(), "{file}");
        let (program, _) = compile(&model, &spec).unwrap();
        let out = program
            .handle(
                &json!({"ranges": ["F!A1:A2", "T!A1"]}),
                &Overrides::default(),
            )
            .unwrap();
        let ranges: &Vec<Json> = out["valueRanges"].as_array().unwrap();
        // `T!A1` holds 1: the formulas double it.
        assert_eq!(ranges[0]["values"], json!([[2], [2]]), "{file}");
        assert_eq!(ranges[1]["values"], json!([[1]]), "{file}");
        // Without ranges there is no default answer.
        assert!(program.handle(&json!({}), &Overrides::default()).is_err());
    }
}

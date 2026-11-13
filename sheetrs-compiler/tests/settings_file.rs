//! Workbooks without a settings sheet: the same declarations in a JSON file named by
//! `sheetc.toml`, and what a workbook exposes when nothing is declared.
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
    loader::load(&sheetrs_io::Book::from_file(asset(name)).unwrap()).unwrap()
}

const SETTINGS: &str = r#"{
  "inputs":  [{ "name": "base", "range": "T!A1" }],
  "outputs": [{ "name": "doubled", "range": "F!A1:A2" }],
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
            .calculate(&json!({"base": 10}), &Overrides::default())
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
    assert!(message(r#"{"inputs": [{"name": "a", "range": "Nowhere!A1"}]}"#).contains("input `a`"));
    assert!(message(r#"{"inputs": ["Nowhere!A1"]}"#).contains("input 1"));
    assert!(
        message(
            r#"{"inputs": [{"name": "a", "range": "T!A1"}], "outputs": [{"name": "A", "range": "T!A2"}]}"#
        )
        .contains("already defined")
    );
    assert!(message(r#"{"inputs": [{"name": "T", "range": "T!A1"}]}"#).contains("also a sheet"));
    assert!(message(r#"{"inputs": ["A1"]}"#).contains("must start with the sheet"));
    assert!(message(r#"{"inputz": []}"#).contains("unknown field"));
    assert!(message(r#"{"inputs": {"a": "T!A1"}}"#).contains("invalid type"));
    assert!(
        message(r#"{"meta": [{"key": "x", "meta_type": "input", "value": "T!A1"}]}"#)
            .contains("under `inputs`")
    );
    // Ranges alone, aliases alone, and both.
    for ok in [
        r#"{"inputs": ["T!A1"], "outputs": ["F!A1"]}"#,
        r#"{"inputs": [{"name": "a", "range": "T!A1", "description": "d"}], "outputs": [{"name": "o", "range": "F!A1"}]}"#,
        r#"{"inputs": ["T!A1"], "outputs": [{"name": "o", "range": "F!A1"}, "F!A2"]}"#,
    ] {
        assert!(spec(ok).is_ok(), "{ok}");
    }
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

/// With nothing declared the formula cells are the outputs and the inputs are the constants they
/// start from; the metadata says what is exposed, and a request names the inputs by range or alias.
#[test]
fn a_workbook_with_no_declaration_exposes_its_formula_cells_and_the_constants_they_start_from() {
    for file in ["formulas.xlsx", "formulas.ods"] {
        let model = model(file);
        let spec = Config::default().spec(&model).unwrap();
        assert!(spec.inputs.is_empty() && spec.outputs.is_empty(), "{file}");
        let (program, _) = compile(&model, &spec).unwrap();
        // The outputs are the rectangles of formula cells; the inputs are the constants.
        assert!(
            program.outputs.iter().any(|o| o.name.starts_with("F!")),
            "{file}"
        );
        assert!(!program.inputs.is_empty(), "{file}");
        let listed = |kind: &str| program.meta.iter().filter(|m| m.meta_type == kind).count();
        assert_eq!(listed("input"), program.inputs.len(), "{file}");
        assert_eq!(listed("output"), program.outputs.len(), "{file}");
        // `T!A1` holds 1 and the formulas double it; it is given by its range.
        let ask = |inputs: Json| {
            program
                .handle(
                    &json!({"inputs": inputs, "ranges": ["F!A1:A2"]}),
                    &Overrides::default(),
                )
                .unwrap()
                .to_json()
        };
        let out = ask(json!({}));
        assert_eq!(out["valueRanges"][0]["values"], json!([[2], [2]]), "{file}");
        // A constant is no output: it is an input, given and not read.
        assert!(
            program
                .handle(&json!({"ranges": ["T!A1"]}), &Overrides::default())
                .is_err(),
            "{file}"
        );
        let out = ask(json!({"T!A1": 10}));
        assert_eq!(
            out["valueRanges"][0]["values"],
            json!([[20], [20]]),
            "{file}"
        );
        // Without ranges the answer is the declared outputs, as value ranges.
        let out = program
            .handle(&json!({}), &Overrides::default())
            .unwrap()
            .to_json();
        assert!(
            out["valueRanges"].as_array().is_some_and(|o| !o.is_empty()),
            "{file}"
        );
    }
}

/// `"inputs": []` declares that there are none: what the outputs read is constant and folds.
#[test]
fn an_empty_list_of_inputs_is_a_declaration_of_none() {
    let model = model("formulas.xlsx");
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("m.json");
    std::fs::write(
        &path,
        r#"{"inputs": [], "outputs": [{"name": "d", "range": "F!A1:A2"}]}"#,
    )
    .unwrap();
    let spec = Config {
        settings_file: Some(path),
        ..Config::default()
    }
    .spec(&model)
    .unwrap();
    assert!(spec.inputs_declared && spec.inputs.is_empty());
    let (program, report) = compile(&model, &spec).unwrap();
    assert!(program.inputs.is_empty());
    assert_eq!(report.formulas_kept, 0, "everything folded");
    assert_eq!(
        program
            .calculate(&Json::Null, &Overrides::default())
            .unwrap(),
        json!({"d": [[2], [2]]})
    );
}

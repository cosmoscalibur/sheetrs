//! The notes of the cells against what `LibreOffice` reports for each file (the `notes` of the
//! oracles made by `make_oracle.py`): the text of every note, on a cell that has more or on one that
//! has nothing, and the cells a search in the notes finds.

use serde_json::{Value, json};
use sheetrs_io::Book;
use sheetrs_protocol::Query;
use sheetrs_query::BookQuery;
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

const FIXTURES: &[(&str, &str)] = &[
    ("cells", "sheetrs-io/tests/assets"),
    ("formulas", "sheetrs-io/tests/assets"),
    ("formats", "sheetrs-io/tests/assets"),
    ("lint_book", "tests"),
    ("book", "sheetrs-edit/tests/assets/structure"),
    ("input", "sheetrs-edit/tests/assets/replace"),
];

fn root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("..")
}

/// `(file, notes)`: what `LibreOffice` says, by sheet and cell (`B2`).
type Notes = BTreeMap<(String, String), String>;

fn oracles() -> Vec<(PathBuf, Notes)> {
    let mut out = Vec::new();
    for (name, folder) in FIXTURES {
        for ext in ["xlsx", "ods"] {
            let oracle = root().join(format!(
                "sheetrs-io/tests/assets/oracle/{name}.{ext}.oracle.json"
            ));
            let json: Value = serde_json::from_slice(&std::fs::read(oracle).unwrap()).unwrap();
            let notes = json["notes"]
                .as_array()
                .unwrap()
                .iter()
                .map(|n| {
                    (
                        (
                            n["sheet"].as_str().unwrap().to_string(),
                            sheetrs_formula::addr::a1(
                                n["row"].as_u64().unwrap() as u32,
                                n["col"].as_u64().unwrap() as u32,
                            ),
                        ),
                        n["text"].as_str().unwrap().to_string(),
                    )
                })
                .collect();
            out.push((root().join(folder).join(format!("{name}.{ext}")), notes));
        }
    }
    out
}

fn read(file: &Path, query: Value) -> Vec<sheetrs_protocol::CellRecord> {
    let query: Query = serde_json::from_value(query).unwrap();
    Book::from_file(file).unwrap().query(&query).unwrap().cells
}

#[test]
fn the_text_of_every_note_is_what_libreoffice_reports() {
    let mut with_notes = 0;
    for (file, expected) in oracles() {
        let got: BTreeMap<_, _> = read(&file, json!({"properties": ["comment"]}))
            .into_iter()
            .filter_map(|c| {
                c.comment
                    .clone()
                    .map(|n| ((c.sheet.to_string(), c.cell()), n))
            })
            .collect();
        assert_eq!(got, expected, "{}", file.display());
        with_notes += usize::from(!expected.is_empty());
    }
    assert!(with_notes >= 4, "the fixtures have notes in both formats");
}

#[test]
fn a_search_in_the_notes_finds_the_cells_whose_note_has_the_text() {
    for (file, expected) in oracles() {
        let want: Vec<_> = expected
            .iter()
            .filter(|(_, text)| text.to_lowercase().contains("hello"))
            .map(|(k, _)| k.clone())
            .collect();
        let got: Vec<_> = read(
            &file,
            json!({"properties": ["comment"], "find": {"text": "hello", "search_in": "comment"}}),
        )
        .into_iter()
        .map(|c| (c.sheet.to_string(), c.cell()))
        .collect();
        let mut want = want;
        want.sort();
        let mut got = got;
        got.sort();
        assert_eq!(got, want, "{}", file.display());
    }
}

#[test]
fn a_note_on_a_cell_with_nothing_else_is_a_cell_with_its_note() {
    for ext in ["xlsx", "ods"] {
        let file = root().join(format!("sheetrs-edit/tests/assets/replace/input.{ext}"));
        let cells = read(
            &file,
            json!({"scope": {"range": "T!D25"}, "properties": ["value", "comment"]}),
        );
        assert_eq!(cells.len(), 1, "{ext}");
        assert_eq!(cells[0].value, Some(Value::Null), "{ext}");
        assert_eq!(
            cells[0].comment.as_deref(),
            Some("Hello on an empty cell"),
            "{ext}"
        );
    }
}

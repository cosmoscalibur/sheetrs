//! The cells of an array formula against `LibreOffice`: it shows every cell of the range as the
//! formula between braces (`{=C14:D14*2}`), and a search in formulas finds all of them.

use serde_json::{Value, json};
use sheetrs_io::Book;
use sheetrs_protocol::Query;
use sheetrs_query::BookQuery;
use std::collections::{BTreeMap, BTreeSet};
use std::path::{Path, PathBuf};

const FIXTURES: &[(&str, &str)] = &[
    ("cells", "sheetrs-io/tests/assets"),
    ("formulas", "sheetrs-io/tests/assets"),
    ("lint_book", "tests"),
    ("book", "sheetrs-edit/tests/assets/structure"),
];

fn root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("..")
}

#[test]
fn every_cell_of_an_array_formula_is_found_by_its_formula_as_libreoffice_shows_it() {
    let mut arrays = 0;
    for (name, folder) in FIXTURES {
        for ext in ["xlsx", "ods"] {
            let oracle = root().join(format!(
                "sheetrs-io/tests/assets/oracle/{name}.{ext}.oracle.json"
            ));
            let json: Value = serde_json::from_slice(&std::fs::read(oracle).unwrap()).unwrap();
            // The cells that `LibreOffice` shows with an array formula, by formula.
            let mut by_formula: BTreeMap<String, BTreeSet<(String, String)>> = BTreeMap::new();
            for cell in json["cells"].as_array().unwrap() {
                if let Some(f) = cell["formula"].as_str().filter(|f| f.starts_with('{')) {
                    by_formula.entry(f.to_string()).or_default().insert((
                        cell["sheet"].as_str().unwrap().to_string(),
                        sheetrs_formula::addr::a1(
                            cell["row"].as_u64().unwrap() as u32,
                            cell["col"].as_u64().unwrap() as u32,
                        ),
                    ));
                }
            }
            let file = root().join(folder).join(format!("{name}.{ext}"));
            for (formula, cells) in by_formula {
                let query: Query = serde_json::from_value(json!({
                    "find": {"text": formula, "search_in": "formula", "entire_cells": true}
                }))
                .unwrap();
                let found: BTreeSet<_> = Book::from_file(&file)
                    .unwrap()
                    .query(&query)
                    .unwrap()
                    .cells
                    .into_iter()
                    .map(|c| (c.sheet.to_string(), c.cell()))
                    .collect();
                assert_eq!(found, cells, "{name}.{ext}: {formula}");
                arrays += 1;
            }
        }
    }
    assert!(
        arrays >= 4,
        "the fixtures have array formulas in both formats"
    );
}

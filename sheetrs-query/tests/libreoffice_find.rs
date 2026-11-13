//! `Find::formatted` is the `SearchFormatted` option of `LibreOffice`: the cells a search finds are
//! those that `findAll` of `LibreOffice` finds in the same file, with and without the option
//! (`tests/assets/find_formatted.json`, made by `tests/assets/make_find_oracle.py`). The searches
//! also cover `Find::filtered` (`SearchFiltered`: rows an autofilter hides, in an ODS and in an XLSX,
//! from workbooks `LibreOffice` saved). The searches include the results of formulas
//! (`include_formulas`), which `LibreOffice` always looks at, and the notes of an XLSX with a
//! threaded comment, of which `LibreOffice` reads the legacy note only.

use sheetrs_io::Book;
use sheetrs_protocol::{Find, Mode, Property, Query, Scope, SearchIn};
use sheetrs_query::BookQuery;
use std::collections::BTreeSet;
use std::path::Path;

#[test]
fn a_search_finds_the_cells_libreoffice_finds_with_and_without_the_formatted_option() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let oracle: Vec<serde_json::Value> = serde_json::from_str(
        &std::fs::read_to_string(root.join("tests/assets/find_formatted.json")).unwrap(),
    )
    .unwrap();
    assert!(!oracle.is_empty());
    let mut bad = Vec::new();
    for case in &oracle {
        let file = case["file"].as_str().unwrap();
        let path = root
            .join("..")
            .join(case["folder"].as_str().unwrap())
            .join(file);
        let source = Book::from_file(&path).unwrap();
        let find = Find {
            mode: if case["regex"].as_bool().unwrap() {
                Mode::Regex
            } else {
                Mode::Literal
            },
            entire_cells: case["entire_cells"].as_bool().unwrap(),
            include_formulas: true,
            formatted: case["formatted"].as_bool().unwrap(),
            filtered: case["filtered"].as_bool().unwrap(),
            search_in: if case["in_notes"].as_bool().unwrap() {
                SearchIn::Comment
            } else {
                SearchIn::Value
            },
            ..Find::literal(case["text"].as_str().unwrap())
        };
        let query = Query {
            scope: vec![Scope::Workbook],
            properties: vec![Property::Value],
            find: Some(find),
            ..Query::default()
        };
        let ours: BTreeSet<String> = source
            .query(&query)
            .unwrap()
            .cells
            .iter()
            .map(|c| format!("{}!{}", c.sheet, c.cell()))
            .collect();
        let libre: BTreeSet<String> = case["found"]
            .as_array()
            .unwrap()
            .iter()
            .map(|v| v.as_str().unwrap().to_string())
            .collect();
        if ours != libre {
            bad.push(format!(
                "{file} {:?} (filtered {}, notes {}, formatted {}): LibreOffice {libre:?}, SheetRS {ours:?}",
                case["text"].as_str().unwrap(),
                case["filtered"],
                case["in_notes"],
                case["formatted"]
            ));
        }
    }
    assert!(
        bad.is_empty(),
        "{} differences:\n{}",
        bad.len(),
        bad.join("\n")
    );
}

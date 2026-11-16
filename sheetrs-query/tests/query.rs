//! The queries over the workbooks `LibreOffice` made in both formats (`sheetrs-io/tests/assets`):
//! every query can be asked of an XLSX and of an ODS, and each answer is what is expected of that
//! file. What `LibreOffice` itself says of each cell is checked in `libreoffice_oracle.rs`.

use serde_json::json;
use sheetrs_io::Book;
use sheetrs_protocol::{Find, Mode, Property, Query, QueryResponse, Reference, Scope, SearchIn};
use sheetrs_query::BookQuery;
use std::path::{Path, PathBuf};

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-io/tests/assets")
        .join(name)
}

fn ask(file: &str, query: &Query) -> QueryResponse {
    Book::from_file(asset(file)).unwrap().query(query).unwrap()
}

fn query(json: serde_json::Value) -> Query {
    serde_json::from_value(json).unwrap()
}

fn both(stem: &str) -> [String; 2] {
    [format!("{stem}.xlsx"), format!("{stem}.ods")]
}

#[test]
fn a_range_gives_the_cells_in_it_and_nothing_else() {
    for file in both("cells") {
        let r = ask(&file, &query(json!({"scope": {"range": "T!A2:A3"}})));
        let cells: Vec<_> = r
            .cells
            .iter()
            .map(sheetrs_protocol::CellRecord::cell)
            .collect();
        assert_eq!(cells, ["A2", "A3"], "{file}");
        assert!(
            r.cells
                .iter()
                .all(|c| &*c.sheet == "T" && c.formula.is_none())
        );
    }
}

#[test]
fn only_the_properties_asked_for_are_in_the_answer() {
    for file in both("cells") {
        let r = ask(
            &file,
            &query(json!({"scope": {"range": "T!A7"}, "properties": ["number_format", "display"]})),
        );
        let c = &r.cells[0];
        assert!(c.value.is_none() && c.formula.is_none(), "{file}");
        assert_eq!(c.display.as_deref(), Some("25.00%"), "{file}");
        assert!(
            c.number_format.as_deref().unwrap().ends_with("0.00%"),
            "{file}"
        );
    }
}

#[test]
fn a_sheet_and_the_workbook_are_read_whole() {
    for file in both("formulas") {
        let sheet = ask(&file, &query(json!({"scope": {"sheet": "F"}})));
        assert_eq!(sheet.cells.len(), 41, "{file}");
        let book = ask(&file, &query(json!({"scope": "workbook"})));
        assert!(book.cells.len() > 41, "{file}");
        let sheets: std::collections::BTreeSet<_> = book.cells.iter().map(|c| &*c.sheet).collect();
        assert!(
            sheets.contains("F") && sheets.contains("T"),
            "{file}: {sheets:?}"
        );
    }
}

#[test]
fn the_default_is_the_values_of_the_workbook() {
    for file in both("formulas") {
        let r = ask(&file, &Query::default());
        assert!(
            r.cells
                .iter()
                .all(|c| c.value.is_some() && c.formula.is_none())
        );
    }
}

#[test]
fn a_sheet_that_does_not_exist_is_an_error_in_both() {
    for file in both("formulas") {
        let source = Book::from_file(asset(&file)).unwrap();
        for scope in [
            Scope::Sheet("Nope".into()),
            Scope::Range("Nope!A1".parse::<Reference>().unwrap()),
        ] {
            let e = source
                .query(&Query {
                    scope: vec![scope],
                    ..Query::default()
                })
                .unwrap_err();
            assert!(e.to_string().contains("no sheet `Nope`"), "{file}: {e}");
        }
    }
}

fn found(file: &str, find: Find) -> Vec<String> {
    ask(
        file,
        &Query {
            find: Some(find),
            ..Query::default()
        },
    )
    .cells
    .iter()
    .map(|c| format!("{}!{}", c.sheet, c.cell()))
    .collect()
}

#[test]
fn a_search_by_value_does_not_look_at_the_results_of_formulas_unless_asked() {
    for file in both("cells") {
        // T!A1 holds a date with an ISO format: the input line writes it as the ISO date.
        let typed = found(&file, Find::literal("2023-03-15"));
        assert!(typed.contains(&"T!A1".to_string()), "{file}: {typed:?}");
        let with = found(
            &file,
            Find {
                include_formulas: true,
                ..Find::literal("2023-03-15")
            },
        );
        assert!(with.len() >= typed.len(), "{file}");
    }
}

#[test]
fn a_search_by_formula_sees_what_libreoffice_shows_for_each_file() {
    for file in both("formulas") {
        // What LibreOffice shows depends on the file: an XLSX has only absolute sheets (`$T.A1`),
        // the ODS of LibreOffice keeps `[T.A1]` as a sheet that is relative (`T.A1`).
        let shown = if std::path::Path::new(&file)
            .extension()
            .is_some_and(|e| e.eq_ignore_ascii_case("ods"))
        {
            "=T.A1*2"
        } else {
            "=$T.A1*2"
        };
        let hits = found(
            &file,
            Find {
                search_in: SearchIn::Formula,
                entire_cells: true,
                ..Find::literal(shown)
            },
        );
        assert!(hits.contains(&"F!A1".to_string()), "{file}: {hits:?}");
        let excel = found(
            &file,
            Find {
                search_in: SearchIn::Formula,
                ..Find::literal("T!A1")
            },
        );
        assert!(
            excel.is_empty(),
            "{file}: Excel's notation is not what is searched"
        );
    }
}

#[test]
fn a_search_by_pattern_and_the_options_of_find_and_replace() {
    for file in both("formulas") {
        let regex = found(
            &file,
            Find {
                mode: Mode::Regex,
                search_in: SearchIn::Formula,
                entire_cells: true,
                ..Find::literal("=SUM\\(.*\\)")
            },
        );
        assert!(!regex.is_empty(), "{file}");
        let wild = found(
            &file,
            Find {
                mode: Mode::Wildcards,
                search_in: SearchIn::Formula,
                entire_cells: true,
                ..Find::literal("=SUM(*)")
            },
        );
        assert_eq!(regex, wild, "{file}");
        let case = found(
            &file,
            Find {
                match_case: true,
                search_in: SearchIn::Formula,
                ..Find::literal("sum(")
            },
        );
        assert!(case.is_empty(), "{file}");
    }
}

#[test]
fn an_unsupported_pattern_is_an_error_and_an_unknown_locale_too() {
    let source = Book::from_file(asset("cells.xlsx")).unwrap();
    let bad = Query {
        find: Some(Find {
            mode: Mode::Regex,
            ..Find::literal("\\bfoo")
        }),
        ..Query::default()
    };
    assert!(source.query(&bad).is_err());
    let locale = Query {
        properties: vec![Property::Display],
        locale: Some("xx-XX".into()),
        ..Query::default()
    };
    assert!(source.query(&locale).is_err());
}

#[test]
fn a_cell_with_no_format_has_general_in_either_file() {
    // An XLSX writes `General`, an ODS writes nothing: LibreOffice reports `General` for both.
    let q = query(json!({"scope": {"range": "T!A2:A3"}, "properties": ["number_format"]}));
    for file in both("formulas") {
        let r = ask(&file, &q);
        assert!(!r.cells.is_empty());
        assert!(
            r.cells
                .iter()
                .all(|c| c.number_format.as_deref() == Some("General")),
            "{file}"
        );
    }
}

#[test]
fn an_error_is_told_from_a_text_that_starts_with_a_hash() {
    for file in both("errors") {
        let r = ask(&file, &query(json!({"scope": {"range": "Errors!A5:A6"}})));
        let values: Vec<_> = r.cells.iter().map(|c| c.value.clone().unwrap()).collect();
        assert_eq!(
            values,
            [json!({"error": "#N/A"}), json!({"error": "#DIV/0!"})],
            "{file}"
        );
    }
}

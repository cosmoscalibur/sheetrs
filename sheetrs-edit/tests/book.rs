//! The editor on an integrated workbook, judged by `LibreOffice` itself.
//!
//! `tests/assets/book/` holds a report workbook made with `LibreOffice` 26.8 in both formats
//! (`book.*`, built by `scripts/gen_book_fixtures.py`: merged cells, a hidden column and row, array
//! formulas, a conditional format, a validation, a hyperlink, a note, a chart, defined names (one
//! local), an external reference, rich text, a hidden sheet) and, for each case of `cases.json`,
//! the workbook `LibreOffice` leaves after doing the same operation itself (`expected_<case>.*`,
//! saved by `LibreOffice` in the same format). The editor runs each case on the book and the two
//! workbooks are read back and compared: the sheets, the names, what each sheet holds of structure
//! (merged ranges, hidden rows and columns, conditional formats, notes, charts, hyperlinks,
//! validations) and every cell, the value and the formula (the cached result of a formula that was
//! rewritten is not compared: the editor drops it and the application computes it). Each file is
//! judged on its own.

use serde_json::Value as Json;
use sheetrs_edit::BookEdit;
use sheetrs_edit::{ColumnRange, RowRange, ValueScope};
use sheetrs_io::Book;
use sheetrs_io::reader::{CellValue, Workbook};
use sheetrs_protocol::{Find, Mode, Replace, Scope, SearchIn};
use sheetrs_query::BookQuery;
use std::collections::BTreeMap;
use std::io::Read;
use std::path::{Path, PathBuf};

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/book")
        .join(name)
}

fn cases() -> Vec<Json> {
    serde_json::from_str(&std::fs::read_to_string(asset("cases.json")).unwrap()).unwrap()
}

fn strings(case: &Json, key: &str) -> Vec<String> {
    case[key]
        .as_array()
        .map(|a| a.iter().map(|s| s.as_str().unwrap().to_string()).collect())
        .unwrap_or_default()
}

/// The editor with the operation a case describes.
fn editor(source: &Book, case: &Json) -> Book {
    let mut editor = source.clone();
    let rows: Vec<RowRange> = strings(case, "delete_rows")
        .iter()
        .map(|s| s.parse().unwrap())
        .collect();
    if !rows.is_empty() {
        editor.delete_rows(rows);
    }
    let columns: Vec<ColumnRange> = strings(case, "delete_columns")
        .iter()
        .map(|s| s.parse().unwrap())
        .collect();
    if !columns.is_empty() {
        editor.delete_columns(columns);
    }
    let values: Vec<ValueScope> = strings(case, "formula_to_value")
        .iter()
        .map(|s| s.parse().unwrap())
        .collect();
    if !values.is_empty() {
        editor.formula_to_value(values);
    }
    let names = strings(case, "delete_names");
    if !names.is_empty() {
        editor.delete_names(names);
    }
    let sheets = strings(case, "delete_sheets");
    if !sheets.is_empty() {
        editor.delete_sheets(sheets);
    }
    if let Some(replace) = case.get("replace") {
        let find = Find {
            text: replace["text"].as_str().unwrap().to_string(),
            mode: Mode::Literal,
            match_case: false,
            entire_cells: false,
            search_in: match replace["search_in"].as_str() {
                Some("formula") => SearchIn::Formula,
                _ => SearchIn::Value,
            },
            include_formulas: false,
            formatted: false,
            filtered: false,
        };
        let scopes: Vec<Scope> = replace["scope"]
            .as_array()
            .map(|a| {
                a.iter()
                    .map(|s| serde_json::from_value(s.clone()).unwrap())
                    .collect()
            })
            .unwrap_or_default();
        editor.replace(
            scopes,
            [Replace {
                find,
                with: replace["with_"].as_str().unwrap().to_string(),
            }],
        );
    }
    editor
}

/// A formula as the file spells it: the reference that went reads `Data!#REF!` in an XLSX and
/// `#REF!` in an ODS, and a range of one cell (`A1:A1`) is that cell.
fn norm(formula: &str) -> String {
    let gone = regex::Regex::new(r"(?:'[^']+'|[A-Za-z0-9_.]+)!#REF!").unwrap();
    let one = regex::Regex::new(r"(\$?[A-Z]+\$?[0-9]+):(\$?[A-Z]+\$?[0-9]+)").unwrap();
    let formula = gone.replace_all(formula, "#REF!");
    one.replace_all(&formula, |c: &regex::Captures| {
        if c[1] == c[2] {
            c[1].to_string()
        } else {
            c[0].to_string()
        }
    })
    .into_owned()
}

type Cells = BTreeMap<(String, u32, u32), (CellValue, Option<String>, bool)>;

fn cells(workbook: &Workbook) -> Cells {
    let mut out = BTreeMap::new();
    for sheet in &workbook.sheets {
        for ((row, col), cell) in sheet.cells.iter() {
            if matches!(cell.value, CellValue::Empty) && cell.formula.is_none() {
                continue;
            }
            // A boolean is the number 1 in one file and `TRUE` in the other.
            let value = match &cell.value {
                CellValue::Boolean(b) => CellValue::Number(f64::from(u8::from(*b))),
                v => v.clone(),
            };
            out.insert(
                (sheet.name.clone(), row, col),
                (value, cell.formula.as_deref().map(norm), cell.is_array),
            );
        }
    }
    out
}

fn same(a: &CellValue, b: &CellValue) -> bool {
    match (a, b) {
        (CellValue::Number(x), CellValue::Number(y)) => (x - y).abs() < 1e-9,
        (a, b) => format!("{a:?}") == format!("{b:?}"),
    }
}

/// The structure of a workbook that the reader reports by sheet.
fn structure(workbook: &Workbook) -> Vec<String> {
    let mut out = Vec::new();
    for sheet in &workbook.sheets {
        out.push(format!(
            "{} {:?} merged={:?} hidden_rows={:?} hidden_columns={:?} cf={} {:?} notes={} charts={} tables={}",
            sheet.name,
            sheet.state,
            sheet.merged_cells,
            sheet.hidden_rows,
            sheet.hidden_columns,
            sheet.conditional_formatting_count,
            sheet.conditional_formatting_ranges,
            sheet.comments,
            sheet.charts,
            sheet.tables,
        ));
    }
    let mut names: Vec<_> = workbook.defined_names.iter().collect();
    names.sort();
    let names: Vec<_> = names.iter().map(|(n, v)| (n.as_str(), norm(v))).collect();
    out.push(format!("names={names:?}"));
    out
}

/// How many hyperlinks and validations a file holds, counted in its XML: the reader does not
/// report them.
fn raw_counts(path: &Path) -> (usize, usize) {
    let mut zip = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
    let mut links = 0;
    let mut validations = 0;
    for i in 0..zip.len() {
        let mut file = zip.by_index(i).unwrap();
        let name = file.name().unwrap().into_owned();
        if !(name.starts_with("xl/worksheets/sheet")
            && std::path::Path::new(&name)
                .extension()
                .is_some_and(|e| e.eq_ignore_ascii_case("xml"))
            || name == "content.xml")
        {
            continue;
        }
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        links += xml.matches("<hyperlink ").count() + xml.matches("<text:a ").count();
        // The declarations of an ODS that no cell uses are not counted: the editor leaves them.
        validations += xml.matches("<dataValidation ").count()
            + xml.matches("table:content-validation-name=").count();
    }
    (links, validations)
}

/// The text of every note of a workbook, by sheet and cell, as `sheetrs query` reads it.
fn notes(path: &Path) -> BTreeMap<(String, String), String> {
    let query: sheetrs_protocol::Query =
        serde_json::from_value(serde_json::json!({"properties": ["comment"]})).unwrap();
    Book::from_file(path)
        .unwrap()
        .query(&query)
        .unwrap()
        .cells
        .into_iter()
        .filter_map(|c| {
            c.comment
                .clone()
                .map(|n| ((c.sheet.to_string(), c.cell()), n))
        })
        .collect()
}

fn check(extension: &str, case: &Json) -> Vec<String> {
    let name = case["name"].as_str().unwrap();
    let dir = tempfile::tempdir().unwrap();
    let source = Book::from_file(asset(&format!("book.{extension}"))).unwrap();
    let out = dir.path().join(format!("{name}.{extension}"));
    let report = match editor(&source, case).write(&out) {
        Ok(report) => report,
        Err(e) => return vec![format!("{extension} {name}: {e}")],
    };
    if !report.skipped.is_empty() {
        return vec![format!("{extension} {name}: {:?}", report.skipped)];
    }
    let expected = asset(&format!("expected_{name}.{extension}"));
    let (got, want) = (
        Book::from_file(&out).unwrap().workbook().unwrap(),
        Book::from_file(&expected).unwrap().workbook().unwrap(),
    );
    let mut different = Vec::new();

    for (g, w) in structure(&got).iter().zip(structure(&want).iter()) {
        if g != w {
            different.push(format!("structure: got {g}\n           LibreOffice {w}"));
        }
    }
    if structure(&got).len() != structure(&want).len() {
        different.push(format!(
            "sheets: got {:?}, LibreOffice {:?}",
            got.sheets.iter().map(|s| &s.name).collect::<Vec<_>>(),
            want.sheets.iter().map(|s| &s.name).collect::<Vec<_>>()
        ));
    }
    let (got_cells, want_cells) = (cells(&got), cells(&want));
    for key in got_cells
        .keys()
        .chain(want_cells.keys())
        .collect::<std::collections::BTreeSet<_>>()
    {
        let equal = match (got_cells.get(key), want_cells.get(key)) {
            (Some((gv, gf, ga)), Some((wv, wf, wa))) => {
                // `LibreOffice` writes a name it no longer has in lower case: the same name.
                let formula = gf == wf
                    || matches!(wv, CellValue::Error(_))
                        && gf.as_ref().map(|f| f.to_lowercase())
                            == wf.as_ref().map(|f| f.to_lowercase());
                formula && ga == wa && (gf.is_some() || same(gv, wv))
            }
            _ => false,
        };
        if !equal {
            different.push(format!(
                "{key:?}: got {:?}, LibreOffice {:?}",
                got_cells.get(key),
                want_cells.get(key)
            ));
        }
    }
    if notes(&out) != notes(&expected) {
        different.push("the notes differ".to_string());
    }
    if raw_counts(&out) != raw_counts(&expected) {
        different.push(format!(
            "hyperlinks and validations: got {:?}, LibreOffice {:?}",
            raw_counts(&out),
            raw_counts(&expected)
        ));
    }
    if different.is_empty() {
        Vec::new()
    } else {
        vec![format!("{extension} {name}:\n  {}", different.join("\n  "))]
    }
}

fn check_all(extension: &str) {
    let problems: Vec<String> = cases()
        .iter()
        .flat_map(|case| check(extension, case))
        .collect();
    assert!(problems.is_empty(), "{}", problems.join("\n"));
}

#[test]
fn an_xlsx_book_is_edited_as_libreoffice_edits_it() {
    check_all("xlsx");
}

#[test]
fn an_ods_book_is_edited_as_libreoffice_edits_it() {
    check_all("ods");
}

/// Every case does something: its expected workbook is not the book.
#[test]
fn every_case_changes_the_book() {
    for extension in ["xlsx", "ods"] {
        let book = Book::from_file(asset(&format!("book.{extension}")))
            .unwrap()
            .workbook()
            .unwrap();
        for case in cases() {
            let name = case["name"].as_str().unwrap();
            let expected = Book::from_file(asset(&format!("expected_{name}.{extension}")))
                .unwrap()
                .workbook()
                .unwrap();
            assert!(
                structure(&book) != structure(&expected) || cells(&book) != cells(&expected),
                "{extension} {name}: LibreOffice left the book as it was"
            );
        }
    }
}

/// Writing results and a constant leaves everything else of the book as it was: the structure, the
/// notes, the hyperlinks and validations, and every other cell.
#[test]
fn writing_cells_keeps_the_rest_of_the_book() {
    use sheetrs_edit::{CachedResult, CellWrites, Constant};
    for extension in ["xlsx", "ods"] {
        let path = asset(&format!("book.{extension}"));
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut writes = CellWrites::new();
        writes
            .constant("Data", 2, 1, Constant::Number(99.0))
            .result("Data", 2, 3, CachedResult::Number(1237.5))
            .result("Data", 2, 4, CachedResult::Text("high".into()));
        let mut editor = Book::from_file(&path).unwrap();
        editor.write_cells(writes);
        assert_eq!(editor.write(&out).unwrap().written, 3, "{extension}");

        let (before, after) = (
            Book::from_file(&path).unwrap().workbook().unwrap(),
            Book::from_file(&out).unwrap().workbook().unwrap(),
        );
        assert_eq!(structure(&before), structure(&after), "{extension}");
        assert_eq!(notes(&path), notes(&out), "{extension}: the notes");
        assert_eq!(raw_counts(&path), raw_counts(&out), "{extension}");
        let (a, b) = (cells(&before), cells(&after));
        let sheet = "Data".to_string();
        assert_eq!(
            b[&(sheet.clone(), 2, 1)].0,
            CellValue::Number(99.0),
            "{extension}"
        );
        assert_eq!(
            b[&(sheet.clone(), 2, 3)].0,
            CellValue::Number(1237.5),
            "{extension}"
        );
        for (key, cell) in &a {
            if key.0 == sheet && key.1 == 2 && matches!(key.2, 1 | 3 | 4) {
                // The formulas stay where results were written.
                if key.2 != 1 {
                    assert_eq!(b[key].1, cell.1, "{extension} {key:?}");
                }
                continue;
            }
            assert!(
                b[key].1 == cell.1
                    && b[key].2 == cell.2
                    && (cell.1.is_some() || same(&b[key].0, &cell.0)),
                "{extension} {key:?}: {:?} became {:?}",
                cell,
                b[key]
            );
        }
    }
}

/// `write_in_place` leaves, in the file itself, what `write` leaves in another.
#[test]
fn writing_in_place_gives_what_libreoffice_gives() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join(format!("book.{extension}"));
        std::fs::copy(asset(&format!("book.{extension}")), &path).unwrap();
        let mut editor = Book::from_file(&path).unwrap();
        editor.delete_columns(["Data!B".parse().unwrap()]);
        editor.write_in_place().unwrap();
        let (got, want) = (
            Book::from_file(&path).unwrap().workbook().unwrap(),
            Book::from_file(asset(&format!("expected_delete_columns.{extension}")))
                .unwrap()
                .workbook()
                .unwrap(),
        );
        assert_eq!(structure(&got), structure(&want), "{extension}");
        let (got, want) = (cells(&got), cells(&want));
        assert_eq!(got.len(), want.len(), "{extension}");
        for (key, (value, formula, array)) in &got {
            let (wvalue, wformula, warray) = &want[key];
            assert!(
                formula == wformula
                    && array == warray
                    && (formula.is_some() || same(value, wvalue)),
                "{extension} {key:?}"
            );
        }
    }
}

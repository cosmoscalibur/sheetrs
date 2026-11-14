//! The in-memory write (`BookEdit::write_to`, what Python and WebAssembly use) against `LibreOffice` 26.8.
//!
//! `tests/assets/write_to/oracle.json` (`make_oracle.py`) holds, for the workbooks `book`, `rules` and `refs` of
//! `tests/assets/use_cache/` in both formats and for each edit of its list, what `LibreOffice` itself leaves
//! when it makes the edit through UNO and calculates every formula (`expected`): the value of each cell of
//! `A1:F80`, the hidden rows and columns and the visibility of each sheet. For each edit this test checks that
//!
//! * the bytes of `write_to` are the bytes of `write`, with the cache and without it;
//! * what `write_to` leaves is what `LibreOffice` left: the values, the hidden rows and columns and the
//!   sheets (an edit that leaves no result stale calculates nothing, and keeps what the file caches);
//! * with `use_cache`, a formula keeps what the file caches or has what the full recalculation gives.
//!
//! The formulas are compared by the script (`differences` of the oracle: what `LibreOffice` shows of the file
//! `sheetrs edit` writes, after calculating it) and are not asserted here: a reference that goes is spelled
//! differently by each writer. The differences the script found are reported to the maintainer, none is
//! approved; `known_differences` below lists the cells of the values that differ, so that a change shows.

use serde_json::Value as Json;
use sheetrs_edit::BookEdit;
use sheetrs_io::reader::{CellValue, SheetState};
use sheetrs_io::{Book, Volatile};
use sheetrs_protocol::{Find, Mode, Replace, SearchIn};
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};
use std::sync::LazyLock;

const ROWS: u32 = 80;
const COLUMNS: u32 = 6;

fn asset(name: &str, extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/use_cache")
        .join(format!("{name}.{extension}"))
}

static ORACLE: LazyLock<Json> = LazyLock::new(|| {
    let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/assets/write_to/oracle.json");
    serde_json::from_str(&std::fs::read_to_string(path).unwrap()).unwrap()
});

/// A cell as the oracle says it: `n` number, `s` text, `e` error.
#[derive(Debug, Clone, PartialEq)]
enum Value {
    Number(f64),
    Text(String),
    Error(String),
}

#[derive(Debug, Default, PartialEq)]
struct Sheet {
    visible: bool,
    hidden_rows: Vec<u32>,
    hidden_columns: Vec<u32>,
    cells: BTreeMap<String, Value>,
}

type Dump = BTreeMap<String, Sheet>;

fn name(row: u32, col: u32) -> String {
    format!(
        "{}{}",
        char::from(b'A' + u8::try_from(col).unwrap()),
        row + 1
    )
}

fn oracle_dump(sheets: &Json) -> Dump {
    sheets
        .as_array()
        .unwrap()
        .iter()
        .map(|s| {
            let ints = |key: &str| -> Vec<u32> {
                s[key]
                    .as_array()
                    .unwrap()
                    .iter()
                    .map(|n| u32::try_from(n.as_u64().unwrap()).unwrap())
                    .collect()
            };
            let cells = s["cells"]
                .as_object()
                .unwrap()
                .iter()
                .map(|(cell, c)| {
                    let v = &c["v"];
                    let value = match v["t"].as_str().unwrap() {
                        "n" => Value::Number(v["v"].as_f64().unwrap()),
                        "e" => Value::Error(v["v"].as_str().unwrap().to_string()),
                        _ => Value::Text(v["v"].as_str().unwrap().to_string()),
                    };
                    (cell.clone(), value)
                })
                .collect();
            let sheet = Sheet {
                visible: s["visible"].as_bool().unwrap(),
                hidden_rows: ints("hidden_rows"),
                hidden_columns: ints("hidden_columns"),
                cells,
            };
            (s["name"].as_str().unwrap().to_string(), sheet)
        })
        .collect()
}

/// The errors an XLSX caches in the cells of each sheet, as the file keeps them. The reader gives `#NAME?` to a
/// cell whose formula holds a bare `#REF!` (`LibreOffice` reads that as a name it does not know and calculates
/// it so, whatever the file caches), and what `LibreOffice` leaves after the edit is what it caches: `#REF!`.
fn cached_errors(bytes: &[u8]) -> BTreeMap<String, BTreeMap<String, String>> {
    use std::io::Read;
    let Ok(mut zip) = zip::ZipArchive::new(std::io::Cursor::new(bytes)) else {
        return BTreeMap::new();
    };
    let mut text = |name: &str| {
        let mut out = String::new();
        zip.by_name(name).ok()?.read_to_string(&mut out).ok()?;
        Some(out)
    };
    let attr = |tag: &str, key: &str| {
        let at = tag.find(&format!(" {key}=\""))? + key.len() + 3;
        Some(tag[at..at + tag[at..].find('"')?].to_string())
    };
    let (Some(rels), Some(workbook)) =
        (text("xl/_rels/workbook.xml.rels"), text("xl/workbook.xml"))
    else {
        return BTreeMap::new(); // an ODS
    };
    let targets: BTreeMap<String, String> = rels
        .split("<Relationship ")
        .skip(1)
        .filter_map(|r| Some((attr(r, "Id")?, attr(r, "Target")?)))
        .collect();
    let cell = regex::Regex::new(
        r#"<c r="([A-Z]+[0-9]+)"[^>]*t="e"[^>]*>(?:<f[^>]*>[^<]*</f>)?<v>([^<]*)</v>"#,
    )
    .unwrap();
    let mut out = BTreeMap::new();
    for sheet in workbook.split("<sheet ").skip(1) {
        let (name, id) = (attr(sheet, "name").unwrap(), attr(sheet, "r:id").unwrap());
        let target = targets[&id].trim_start_matches('/');
        let xml = text(&format!("xl/{}", target.trim_start_matches("xl/"))).unwrap();
        out.insert(
            name,
            cell.captures_iter(&xml)
                .map(|c| (c[1].to_string(), c[2].to_string()))
                .collect(),
        );
    }
    out
}

/// What the workbook `bytes` holds of the area the oracle compares.
fn dump(bytes: &[u8]) -> Dump {
    let workbook = Book::from_bytes(bytes.to_vec())
        .unwrap()
        .workbook()
        .unwrap();
    let cached = cached_errors(bytes);
    workbook
        .sheets
        .iter()
        .map(|sheet| {
            let mut cells = BTreeMap::new();
            for ((row, col), cell) in &sheet.cells {
                if row >= ROWS || col >= COLUMNS {
                    continue;
                }
                let raw = cached
                    .get(&sheet.name)
                    .and_then(|errors| errors.get(&name(row, col)));
                let value = match &cell.value {
                    CellValue::Empty => continue,
                    CellValue::Error(_) if raw.is_some() => Value::Error(raw.unwrap().clone()),
                    CellValue::Number(n) => Value::Number(*n),
                    CellValue::Boolean(b) => Value::Number(f64::from(u8::from(*b))),
                    CellValue::Text(t) => Value::Text(t.to_string()),
                    CellValue::Error(e) => Value::Error(e.to_string()),
                };
                cells.insert(name(row, col), value);
            }
            let hidden = |ranges: &[(u32, u32)], limit: u32| -> Vec<u32> {
                (0..limit)
                    .filter(|i| ranges.iter().any(|(a, b)| (*a..=*b).contains(i)))
                    .collect()
            };
            let one = Sheet {
                visible: sheet.state == SheetState::Visible,
                hidden_rows: hidden(&sheet.hidden_rows, ROWS),
                hidden_columns: hidden(&sheet.hidden_columns, COLUMNS),
                cells,
            };
            (sheet.name.clone(), one)
        })
        .collect()
}

fn same_value(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Number(x), Value::Number(y)) => (x - y).abs() <= 1e-9 * x.abs().max(1.0),
        _ => a == b,
    }
}

/// The edit as the oracle lists it, on the methods of `BookEdit`.
fn apply(book: &mut Book, operations: &Json) {
    for operation in operations.as_array().unwrap() {
        let operation = operation.as_array().unwrap();
        let kind = operation[0].as_str().unwrap();
        let text = || operation[1].as_str().unwrap().to_string();
        match kind {
            "delete_rows" => drop(book.delete_rows([text().parse().unwrap()])),
            "hide_rows" => drop(book.hide_rows([text().parse().unwrap()])),
            "show_rows" => drop(book.show_rows([text().parse().unwrap()])),
            "delete_columns" => drop(book.delete_columns([text().parse().unwrap()])),
            "hide_columns" => drop(book.hide_columns([text().parse().unwrap()])),
            "show_columns" => drop(book.show_columns([text().parse().unwrap()])),
            "delete_sheets" => drop(book.delete_sheets([text()])),
            "hide_sheets" => drop(book.hide_sheets([text()])),
            "show_sheets" => drop(book.show_sheets([text()])),
            "delete_names" => drop(book.delete_names([text()])),
            "formula_to_value" => drop(book.formula_to_value([text().parse().unwrap()])),
            "replace" => {
                let spec = &operation[1];
                let find = Find {
                    mode: if spec["regex"].as_bool().unwrap_or(false) {
                        Mode::Regex
                    } else {
                        Mode::Literal
                    },
                    search_in: if spec["in"].as_str() == Some("formula") {
                        SearchIn::Formula
                    } else {
                        SearchIn::Value
                    },
                    ..Find::literal(spec["text"].as_str().unwrap())
                };
                book.replace(
                    [spec["scope"].as_str().unwrap().into()],
                    [Replace {
                        find,
                        with: spec["with"].as_str().unwrap().to_string(),
                    }],
                );
            }
            other => panic!("{other}"),
        }
    }
}

fn written(source: &Path, operations: &Json, use_cache: bool) -> (Vec<u8>, bool) {
    let extension = source.extension().unwrap().to_str().unwrap();
    let mut book = Book::from_file(source)
        .unwrap()
        .with_use_cache(use_cache)
        .with_volatile(Volatile::default());
    apply(&mut book, operations);
    let recalculates = book.explain().contains("recalculate every formula");
    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join(format!("out.{extension}"));
    book.write(&out).unwrap();
    let mut bytes = Vec::new();
    book.write_to(&mut bytes).unwrap();
    assert_eq!(
        bytes,
        std::fs::read(&out).unwrap(),
        "write_to is write, use_cache {use_cache}"
    );
    (bytes, recalculates)
}

/// The cells where the values of `got` are not those of `expected` (a sheet's visibility and hidden rows and
/// columns are compared on their own).
fn value_differences(got: &Dump, expected: &Dump) -> Vec<String> {
    let mut out = Vec::new();
    for (sheet, e) in expected {
        let Some(g) = got.get(sheet) else {
            out.push(format!("{sheet}: no sheet"));
            continue;
        };
        for cell in e.cells.keys().chain(g.cells.keys()) {
            let ok = match (e.cells.get(cell), g.cells.get(cell)) {
                (Some(x), Some(y)) => same_value(x, y),
                (None, None) => true,
                _ => false,
            };
            let line = format!(
                "{sheet}!{cell}: {:?} -> {:?}",
                e.cells.get(cell),
                g.cells.get(cell)
            );
            if !ok && !out.contains(&line) {
                out.push(line);
            }
        }
    }
    for sheet in got.keys().filter(|s| !expected.contains_key(*s)) {
        out.push(format!("{sheet}: a sheet that is not expected"));
    }
    out
}

fn layout_differences(got: &Dump, expected: &Dump) -> Vec<String> {
    let mut out = Vec::new();
    for (sheet, e) in expected {
        if let Some(g) = got.get(sheet) {
            if g.visible != e.visible {
                out.push(format!("{sheet} visible: {} -> {}", e.visible, g.visible));
            }
            if g.hidden_rows != e.hidden_rows {
                out.push(format!(
                    "{sheet} rows: {:?} -> {:?}",
                    e.hidden_rows, g.hidden_rows
                ));
            }
            if g.hidden_columns != e.hidden_columns {
                out.push(format!(
                    "{sheet} columns: {:?} -> {:?}",
                    e.hidden_columns, g.hidden_columns
                ));
            }
        }
    }
    out
}

#[test]
fn the_in_memory_write_is_the_file_write_and_gives_what_libreoffice_gives() {
    let mut found: Vec<String> = Vec::new();
    for (fixture, by_format) in ORACLE["fixtures"].as_object().unwrap() {
        for (extension, entry) in by_format.as_object().unwrap() {
            let source = asset(fixture, extension);
            // What the file caches, as the reader reads it (`LibreOffice` calculates some formulas when it loads an ODS).
            let cached = dump(&std::fs::read(&source).unwrap());
            for edit in entry["edits"].as_array().unwrap() {
                let label = edit["label"].as_str().unwrap();
                let context = format!("{fixture}.{extension}: {label}");
                let expected = oracle_dump(&edit["expected"]);
                let operations = &edit["operations"];
                let (whole, recalculates) = written(&source, operations, false);
                let (kept, _) = written(&source, operations, true);
                let (got, with_cache) = (dump(&whole), dump(&kept));

                // The hidden rows and columns and the sheets: what `LibreOffice` leaves.
                assert_eq!(
                    layout_differences(&got, &expected),
                    Vec::<String>::new(),
                    "{context}"
                );
                assert_eq!(
                    layout_differences(&with_cache, &expected),
                    Vec::<String>::new(),
                    "{context}"
                );

                if recalculates {
                    let lines = value_differences(&got, &expected);
                    if let Some(first) = lines.first() {
                        let cell = first.split(':').next().unwrap();
                        found.push(format!("{context}: {} cells, first {cell}", lines.len()));
                    }
                } else {
                    // Nothing was left stale: the cells are the ones the file caches, which no edit calculated.
                    for (sheet, c) in &cached {
                        let g = &got[sheet];
                        for (cell, value) in &c.cells {
                            if sheet == "Calc" {
                                assert!(
                                    same_value(value, &g.cells[cell]),
                                    "{context}: {sheet}!{cell}"
                                );
                            }
                        }
                    }
                    assert_eq!(
                        got, with_cache,
                        "{context}: nothing to calculate, nothing to keep"
                    );
                }

                // With the cache a formula keeps what the file caches, or has what the full run gives.
                for (sheet, g) in &got {
                    if sheet != "Calc" {
                        assert_eq!(g, &with_cache[sheet], "{context}: {sheet}");
                        continue;
                    }
                    let u = &with_cache[sheet];
                    for (cell, value) in &g.cells {
                        let kept = &u.cells[cell];
                        let was = cached.get(sheet).and_then(|c| c.cells.get(cell));
                        assert!(
                            same_value(kept, value) || was.is_some_and(|w| same_value(kept, w)),
                            "{context}: {sheet}!{cell} is {kept:?}, the full run gives {value:?} and the file caches {was:?}"
                        );
                    }
                }
            }
        }
    }
    assert_eq!(
        found,
        Vec::<String>::new(),
        "values that differ from LibreOffice's (none is approved)"
    );
}

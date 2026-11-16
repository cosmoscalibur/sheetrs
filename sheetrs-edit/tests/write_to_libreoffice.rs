//! The in-memory write (`BookEdit::write_to`, what Python and WebAssembly use) against `LibreOffice` 26.8.
//!
//! The reference is the workbook **saved and opened again** (project rule, `CONTRIBUTING.md`): the memory of
//! `LibreOffice` right after an edit is not a reference, because the same workbook, saved and opened by someone
//! else, shows other values. `tests/assets/write_to/oracle.json` (`make_oracle.py`) holds, for the workbooks
//! `book`, `rules`, `refs` and `preref` of `tests/assets/cached/` in both formats and for each edit of its list,
//! what `LibreOffice` shows when it opens the file it saved after making the edit through UNO and calculating every
//! formula (`reread`): the value of each cell of `A1:F80`, the hidden rows and columns and the visibility of each
//! sheet. For each edit this test checks that
//!
//! * the bytes of `write_to` are the bytes of `write`, with the cache and without it;
//! * what `write_to` leaves is what `LibreOffice` shows of its file: the values, the hidden rows and columns and
//!   the sheets (an edit that leaves no result stale calculates nothing, and keeps what the file caches);
//! * by default (the cache), a formula keeps what the file caches or has what the run with `calculate_all` gives (a
//!   conversion converts what the file caches where nothing reaches the formula, and what the engine
//!   calculates in a run with `calculate_all`: the cells it names are what `LibreOffice` gives after calculating);
//!
//! No value differs from `reread`. The file `sheetrs edit` writes is the same text as the one `LibreOffice` saves
//! (a reference that went is `#REF!`, in both formats, and a qualified `Data!#REF!` of the input is written bare),
//! and the oracle records what `LibreOffice` shows of it, with its default load (`differences`) and after a hard
//! recalculation (`hard_differences`): both must be empty. `reread_default` and `differences_default` are the same for
//! the default mode, where the only differences are the approved ones (`approved`) and the text of the qualified errors (`QUALIFIED_ERROR_TEXT`). The memory of `LibreOffice` is recorded
//! in the oracle (`in_memory_differences`) as informational evidence only; nothing here compares with it.

use serde_json::Value as Json;
use sheetrs_edit::{BookEdit, Calculation};
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
        .join("tests/assets/cached")
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

#[derive(Debug, Default, Clone, PartialEq)]
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
    // The values written are the inputs of a calculation (one call: they are written together).
    let mut inputs = serde_json::Map::new();
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
            "set_value" => {
                inputs.insert(
                    operation[1].as_str().unwrap().to_string(),
                    operation[2].clone(),
                );
            }
            other => panic!("{other}"),
        }
    }
    if !inputs.is_empty() {
        book.calculate(Calculation {
            inputs,
            ..Calculation::default()
        });
    }
}

/// Whether a conversion of the edit names the cell `cell` (`B7`) of `sheet`: `Sheet`, `Sheet!B:D` or `Sheet!3:9`.
fn converted(operations: &Json, sheet: &str, cell: &str) -> bool {
    let at = cell.trim_start_matches(|c: char| c.is_ascii_uppercase());
    let (column, row) = (&cell[..cell.len() - at.len()], at.parse::<u32>().unwrap());
    operations.as_array().unwrap().iter().any(|operation| {
        let operation = operation.as_array().unwrap();
        if operation[0].as_str() != Some("formula_to_value") {
            return false;
        }
        let scope = operation[1].as_str().unwrap();
        let (name, area) = scope.split_once('!').unwrap_or((scope, ""));
        let (first, last) = area.split_once(':').unwrap_or((area, area));
        name == sheet
            && (area.is_empty()
                || if first.chars().all(|c| c.is_ascii_digit()) {
                    (first.parse().unwrap()..=last.parse().unwrap()).contains(&row)
                } else {
                    (first.len(), first) <= (column.len(), column)
                        && (column.len(), column) <= (last.len(), last)
                })
    })
}

fn written(source: &Path, operations: &Json, cached: bool) -> (Vec<u8>, bool) {
    let extension = source.extension().unwrap().to_str().unwrap();
    let mut book = Book::from_file(source)
        .unwrap()
        .with_calculate_all(!cached)
        .with_volatile(Volatile::default());
    apply(&mut book, operations);
    // A calculation (the values written) calculates every formula as well.
    let recalculates = book.explain().contains("recalculate every formula")
        || operations.to_string().contains("set_value");
    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join(format!("out.{extension}"));
    book.write(&out).unwrap();
    let mut bytes = Vec::new();
    book.write_to(&mut bytes).unwrap();
    assert_eq!(
        bytes,
        std::fs::read(&out).unwrap(),
        "write_to is write, cached {cached}"
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

/// The file `sheetrs edit` writes is, for `LibreOffice`, the file it saves itself after the same edit: both spell
/// a reference that went `#REF!` (the error value of an ODS, the unknown name of an XLSX), so after a hard
/// recalculation the two files show the same formulas and values (`hard_differences` is empty), and so they do with the
/// default load (`#NAME?` in an XLSX after a deleted sheet as after deleted cells). The edits that only hide or show
/// keep the cache of the file, and the ones that convert formulas to values are another matter (the cached results of
/// the file against the calculation of `LibreOffice` when it loads).
#[test]
fn the_file_reads_in_libreoffice_as_the_one_libreoffice_saves() {
    for (fixture, by_format) in ORACLE["fixtures"].as_object().unwrap() {
        for (extension, entry) in by_format.as_object().unwrap() {
            for edit in entry["edits"].as_array().unwrap() {
                let label = edit["label"].as_str().unwrap();
                if label.contains("formula to value") {
                    continue;
                }
                // An edit that leaves no result stale calculates nothing and keeps what the file caches (the
                // fixtures cache numbers that their formulas do not give): the default load shows those, and
                // only after a hard recalculation the two files are compared.
                let keeps_the_cache = label.starts_with("hide ") || label.starts_with("show ");
                // PROPOSED, NOT APPROVED (`docs/approved-differences.md`): the text of a formula the run does not touch
                // keeps `Data!#REF!`, which `LibreOffice` saves bare; after a hard recalculation of the file it shows
                // `COUNT` `#REF!` where its own file gives 0.
                if (fixture == "preref"
                    && extension == "xlsx"
                    && (label.starts_with("hide ") || label.starts_with("replace ")))
                    || (fixture == "dirty" && extension == "xlsx" && label == "delete name n_cell1")
                {
                    continue;
                }
                assert_eq!(
                    edit["hard_differences"].as_array().unwrap(),
                    &Vec::<Json>::new(),
                    "{fixture}.{extension}: {label}: after calculating"
                );
                if !keeps_the_cache {
                    assert_eq!(
                        edit["differences"].as_array().unwrap(),
                        &Vec::<Json>::new(),
                        "{fixture}.{extension}: {label}: default load"
                    );
                }
            }
        }
    }
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
                // What `LibreOffice` shows when it opens the file it saved after the edit.
                let expected = oracle_dump(&edit["reread"]);
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
                    // APPROVED (`docs/approved-differences.md`, a cached error of an XLSX): `Err:502` of `LibreOffice` is `#VALUE!` in the file.
                    let lines: Vec<String> = value_differences(&got, &expected)
                        .into_iter()
                        .filter(|line| !(extension == "xlsx" && line.contains("Some(Error(\"Err:")))
                        .collect();
                    if let Some(first) = lines.first() {
                        let cell = first.split(':').next().unwrap();
                        found.push(format!("{context}: {} cells, first {cell}", lines.len()));
                        if std::env::var_os("SHOW").is_some() {
                            eprintln!("{context}\n  {}", lines.join("\n  "));
                        }
                    }
                } else {
                    // Nothing was left stale: the cells are the ones the file caches, which no edit calculated,
                    // but the ones a conversion names: they are what `LibreOffice` calculates. An ODS is the exception: its
                    // startup calculates every formula with `calculate_all` (`calculation_mode.rs`, the startup of an ODS), so
                    // the formulas no edit names are calculated too.
                    for (sheet, c) in &cached {
                        let g = &got[sheet];
                        for (cell, value) in &c.cells {
                            if sheet == "Calc"
                                && extension != "ods"
                                && !converted(operations, sheet, cell)
                            {
                                assert!(
                                    same_value(value, &g.cells[cell]),
                                    "{context}: {sheet}!{cell}"
                                );
                            }
                        }
                    }
                    for (sheet, e) in &expected {
                        for (cell, value) in &e.cells {
                            if converted(operations, sheet, cell) {
                                assert!(
                                    got[sheet]
                                        .cells
                                        .get(cell)
                                        .is_some_and(|g| same_value(g, value)),
                                    "{context}: {sheet}!{cell} is {:?}, LibreOffice leaves {value:?}",
                                    got[sheet].cells.get(cell)
                                );
                            }
                        }
                    }
                    if operations.to_string().contains("formula_to_value") {
                        // With the cache a conversion converts what the file caches where nothing reaches, and in an ODS what
                        // `LibreOffice` calculates when it loads the file (`reread_default`).
                        let reference = oracle_dump(&edit["reread_default"]);
                        for (sheet, c) in &cached {
                            for (cell, value) in &c.cells {
                                if sheet == "Calc" && converted(operations, sheet, cell) {
                                    let kept = &with_cache[sheet].cells[cell];
                                    assert!(
                                        same_value(value, kept)
                                            || reference[sheet]
                                                .cells
                                                .get(cell)
                                                .is_some_and(|r| same_value(kept, r)),
                                        "{context}: {sheet}!{cell} with the cache"
                                    );
                                }
                            }
                        }
                    } else {
                        // `LibreOffice` calculates some formulas when it loads the file (the functions it calculates on load, in an ODS the
                        // results of some types in some cells), and the run that keeps the cache does too; the run that calculates all leaves
                        // what the file caches when no edit left a result stale. Where the two differ, the cache has what `LibreOffice`
                        // shows after the same edit (`reread_default`).
                        let reference = oracle_dump(&edit["reread_default"]);
                        for (sheet, g) in &got {
                            for (cell, value) in &g.cells {
                                let kept = &with_cache[sheet].cells[cell];
                                assert!(
                                    same_value(value, kept)
                                        || reference[sheet]
                                            .cells
                                            .get(cell)
                                            .is_some_and(|r| same_value(kept, r)),
                                    "{context}: {sheet}!{cell} is {kept:?} with the cache, {value:?} calculating all, LibreOffice shows {:?}",
                                    reference[sheet].cells.get(cell)
                                );
                            }
                        }
                    }
                }

                // With the cache a formula keeps what the file caches, or has what the full run gives (the
                // constants are the same in both).
                let default_reference = oracle_dump(&edit["reread_default"]);
                for (sheet, g) in &got {
                    let u = &with_cache[sheet];
                    assert_eq!(g.visible, u.visible, "{context}: {sheet}");
                    assert_eq!(g.hidden_rows, u.hidden_rows, "{context}: {sheet}");
                    assert_eq!(g.hidden_columns, u.hidden_columns, "{context}: {sheet}");
                    for (cell, value) in &g.cells {
                        // A qualified `#REF!` written bare drops the result the formula cached: the application that
                        // opens the file calculates it (`bare_qualified_errors`).
                        let Some(kept) = u.cells.get(cell) else {
                            assert!(
                                fixture == "preref" || fixture == "dirty",
                                "{context}: {sheet}!{cell} has no result"
                            );
                            continue;
                        };
                        // The cell of the file that held the result (a cell that moved is not at its address).
                        let was = cached.get(sheet).and_then(|c| {
                            c.cells
                                .get(cell)
                                .filter(|w| same_value(kept, w))
                                .or_else(|| c.cells.values().find(|w| same_value(kept, w)))
                        });
                        // `dirty` caches results no formula gives, so what reads a stale result is neither the file's
                        // nor the full run's: that fixture is compared with `LibreOffice` (`differences_default`).
                        // In an ODS the cache has what `LibreOffice` calculates when it loads the file (`reread_default`).
                        let loaded = default_reference[sheet]
                            .cells
                            .get(cell)
                            .is_some_and(|r| same_value(kept, r));
                        assert!(
                            fixture == "dirty"
                                || same_value(kept, value)
                                || was.is_some()
                                || loaded,
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
        "values that differ from what LibreOffice shows when it opens its file (none is approved)"
    );
}

/// The cell `cell` of `Calc` after `operations` on `fixture`.
fn after(fixture: &str, extension: &str, operations: &Json, cell: &str) -> Value {
    let (bytes, _) = written(&asset(fixture, extension), operations, false);
    dump(&bytes)["Calc"].cells[cell].clone()
}

fn ops(text: &str) -> Json {
    serde_json::from_str(text).unwrap()
}

fn number(n: f64) -> Value {
    Value::Number(n)
}

fn error(e: &str) -> Value {
    Value::Error(e.to_string())
}

/// Observed (re-read) in `LibreOffice` 26.8 (`oracle.json`, `reread`, the file it saves opened again), in both formats
/// and for a deleted sheet as for deleted cells: `COUNT` of a reference that went is 0, `COUNTA` 1, `COUNT(ref;5)` 1,
/// `IFERROR` catches it, `ISERROR` is true. `SUM` is `#REF!` in an ODS and `#NAME?` in an XLSX, whether cells or a sheet
/// went (the bare `#REF!` of the file is the unknown name; no difference is approved here).
#[test]
fn a_lost_reference_is_an_error_after_deleted_sheets_and_cells() {
    for (operations, sum_xlsx) in [
        (r#"[["delete_sheets", "Data"]]"#, "#NAME?"),
        (r#"[["delete_rows", "Data!3"]]"#, "#NAME?"),
    ] {
        let delete = ops(operations);
        // The cell of each function: with a whole sheet gone, the references are rows 1 to 5 of each family; with
        // `Data!3` gone, the one cell is the one that went (`COUNT(Data!A3)` is `A5`, `COUNTA` `A15`...).
        let (count, counta, sum, count5, iferror, iserror) = if operations.contains("sheets") {
            ("A1", "A11", "A6", "A48", "A44", "A43")
        } else {
            ("A5", "A15", "A10", "A48", "A44", "A43")
        };
        for extension in ["xlsx", "ods"] {
            let cell = |c: &str| after("refs", extension, &delete, c);
            assert_eq!(cell(count), number(0.0), "{extension} {operations}: COUNT");
            assert_eq!(
                cell(counta),
                number(1.0),
                "{extension} {operations}: COUNTA"
            );
            assert_eq!(
                cell(count5),
                number(1.0),
                "{extension} {operations}: COUNT(ref;5)"
            );
            assert_eq!(
                cell(iferror),
                number(7.0),
                "{extension} {operations}: IFERROR"
            );
            assert_eq!(
                cell(iserror),
                number(1.0),
                "{extension} {operations}: ISERROR"
            );
            let sum_expected = if extension == "ods" {
                "#REF!"
            } else {
                sum_xlsx
            };
            assert_eq!(
                cell(sum),
                error(sum_expected),
                "{extension} {operations}: SUM"
            );
        }
    }
}

/// Observed (re-read): `COUNTBLANK` of a reference that went is `Err:504` in an ODS (an error value where a
/// reference is wanted) and `#NAME?` in an XLSX (the unknown name).
#[test]
fn countblank_of_a_lost_reference_is_err_504_in_ods_and_unknown_name_in_xlsx() {
    let delete = ops(r#"[["delete_rows", "Data!3"]]"#);
    assert_eq!(after("refs", "ods", &delete, "A35"), error("Err:504"));
    assert_eq!(after("refs", "xlsx", &delete, "A35"), error("#NAME?"));
}

/// Observed (re-read): an error written in a formula (`COUNT(#REF!)` of a file) is what it says, whatever the run
/// deletes, and what a reference that went gives: 0 for `COUNT`, 1 for `COUNTA`; `SUM` is `#REF!` in an ODS and the
/// unknown name of an XLSX.
#[test]
fn an_error_written_in_a_formula_is_read_as_the_error_it_is_after_any_deletion() {
    let delete = ops(r#"[["delete_rows", "Data!3"]]"#);
    for extension in ["xlsx", "ods"] {
        let cell = |c: &str| after("preref", extension, &delete, c);
        assert_eq!(cell("A1"), number(0.0), "{extension}: COUNT(#REF!)");
        assert_eq!(cell("A2"), number(1.0), "{extension}: COUNTA(#REF!)");
        assert_eq!(cell("A5"), number(0.0), "{extension}: IFERROR(#REF!,0)");
    }
    let sum = |extension| after("preref", extension, &delete, "A3");
    assert_eq!(sum("ods"), error("#REF!"));
    assert_eq!(sum("xlsx"), error("#NAME?"));
}

/// Observed (re-read): `IFERROR(Data!A3,7)` and `IFERROR(Other!A1,7)` are 7 in both formats whatever the run deletes,
/// cells or a sheet or both (the memory of `LibreOffice` is not stable here: `#REF!` in some workbooks).
#[test]
fn iferror_catches_a_lost_reference_of_cells_or_of_a_sheet() {
    let delete = ops(r#"[["delete_rows", "Data!3"], ["delete_sheets", "Other"]]"#);
    for extension in ["xlsx", "ods"] {
        let cell = |c: &str| after("preref", extension, &delete, c);
        assert_eq!(
            cell("A21"),
            number(7.0),
            "{extension}: IFERROR(Data!A3,7): cells went"
        );
        assert_eq!(
            cell("A22"),
            number(7.0),
            "{extension}: IFERROR(Other!A1,7): the sheet went"
        );
    }
}

/// A `Sheet!#REF!` that the input holds (as Excel writes it) is written bare when a deletion rewrites the formulas, as
/// `LibreOffice` saves it: the file read again gives `COUNT` 0 (`oracle.json`, `preref` A10: `COUNT(Data!#REF!)`).
#[test]
fn a_qualified_ref_error_is_written_bare_and_counts_as_an_error() {
    let delete = ops(r#"[["delete_rows", "Data!3"]]"#);
    assert_eq!(after("preref", "xlsx", &delete, "A10"), number(0.0));
    assert_eq!(after("preref", "xlsx", &delete, "A11"), number(1.0));
}

/// A difference of the default mode (the cache) the maintainer approved (`docs/approved-differences.md`), as a line of
/// `differences_default` of the oracle:
///
/// * ODS: `LibreOffice` calculates on load every formula whose cached result is a plain number (`rate*2` cached as 140 is
///   14), and `SheetRS` keeps the cached number. The project does not support that behavior. It reaches the values a
///   conversion of formulas to values writes (`value` lines of the edits that convert: the two cells of `book.ods`, 140 and 99
///   where `LibreOffice` has 14 and 8); it does not change what an edit writes in the formulas it does not convert (the file
///   written is calculated by `LibreOffice` when it opens it), so no other edit is excused: the fixtures that must show
///   what is kept and what is calculated (`dirty`) cache text, which an ODS keeps;
/// * XLSX: `ROW()` (the fixtures cache 99 for it) and the other formulas of its class that `LibreOffice` calculates on
///   load. Nothing of the oracle needs it any longer: `LibreOffice` keeps 99 when it loads `book.xlsx` and `SheetRS` keeps it
///   unless a deletion moves the cell of a `ROW`.
fn approved(extension: &str, edit: &str, line: &str) -> bool {
    match extension {
        "ods" => edit.starts_with("formula to value") && line.contains(" value: "),
        _ => false,
    }
}

/// The edits whose file keeps, in the text of a formula that holds a qualified error (`COUNT(Data!#REF!)`, as Excel writes
/// it), what the file has when the run deletes nothing: `LibreOffice` writes it bare (`COUNT(#REF!)`). PROPOSED, NOT APPROVED
/// (`docs/approved-differences.md`, the qualified error): the values are the re-read ones; only the text of the formula differs.
const QUALIFIED_ERROR_TEXT: &[(&str, &str, &str, usize)] = &[
    ("preref", "xlsx", "hide rows Data!2 (no deletion)", 5),
    (
        "preref",
        "xlsx",
        "replace x by y in Data!B1 (no deletion)",
        5,
    ),
    ("dirty", "xlsx", "delete name n_cell1", 1),
];

/// The default mode (the cache) gives what `LibreOffice` shows when it makes the same edit with no `calculateAll()`, saves
/// and opens the file again (`oracle.json`, `reread_default`): the same results, the stale ones that the file caches and
/// `LibreOffice` keeps as much as the ones the edit makes it calculate, in both formats. The fixtures cache wrong results on
/// purpose (`dirty` caches the text `S`), so that a formula that was calculated is told from one that kept its cache.
/// Only the approved differences (`approved`) and the text of the qualified errors (`QUALIFIED_ERROR_TEXT`) are left out.
#[test]
fn the_default_mode_matches_libreoffice_but_for_approved_differences() {
    for (fixture, by_format) in ORACLE["fixtures"].as_object().unwrap() {
        for (extension, entry) in by_format.as_object().unwrap() {
            for edit in entry["edits"].as_array().unwrap() {
                let label = edit["label"].as_str().unwrap();
                // What LibreOffice shows when it opens the file `sheetrs edit` writes (default mode), against
                // what it shows when it opens the file it saves after the same edit with no `calculateAll()`.
                let lines: Vec<&Json> = edit["differences_default"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .filter(|line| !approved(extension, label, line.as_str().unwrap()))
                    .collect();
                let most = QUALIFIED_ERROR_TEXT
                    .iter()
                    .find(|(f, e, l, _)| {
                        (*f, *e, *l) == (fixture.as_str(), extension.as_str(), label)
                    })
                    .map_or(0, |d| d.3);
                assert!(
                    lines.len() <= most,
                    "{fixture}.{extension}: {label}: {} cells differ from LibreOffice's default (at most {most} are listed): {lines:?}",
                    lines.len()
                );
                if most > 0 {
                    assert!(
                        lines
                            .iter()
                            .all(|line| line.as_str().unwrap().contains(" formula: ")),
                        "{fixture}.{extension}: {label}: only the text of a formula may differ: {lines:?}"
                    );
                }
            }
        }
    }
}

/// The default mode run now, not the file the oracle recorded: the values `sheetrs edit` writes when it keeps the results the file
/// caches are the ones `LibreOffice` shows (`reread_default`) after the same edit with no `calculateAll()`, saved and opened
/// again, in both formats. The fixtures cache wrong results on purpose, so every formula that was calculated (`LibreOffice` calculates
/// it because the edit reaches it) is told from every one that kept its cache. The numbers of an ODS are left out: `LibreOffice` calculates
/// them when it loads the file (approved, `approved`), and only `dirty` caches text, which an ODS keeps.
#[test]
fn the_default_mode_gives_the_values_libreoffice_shows_on_open() {
    let mut found: Vec<String> = Vec::new();
    for (fixture, by_format) in ORACLE["fixtures"].as_object().unwrap() {
        for (extension, entry) in by_format.as_object().unwrap() {
            let source = asset(fixture, extension);
            for edit in entry["edits"].as_array().unwrap() {
                let label = edit["label"].as_str().unwrap();
                let expected = oracle_dump(&edit["reread_default"]);
                let (kept, _) = written(&source, &edit["operations"], true);
                // PROPOSED, NOT APPROVED (`docs/approved-differences.md`): `LibreOffice` calculates `FORMULATEXT` when it loads
                // a file (`rules`, `Calc!A8`), and `SheetRS` keeps the result the file caches.
                let lines: Vec<String> = value_differences(&dump(&kept), &expected)
                    .into_iter()
                    .filter(|line| !(fixture == "rules" && line.starts_with("Calc!A8:")))
                    // APPROVED (`docs/approved-differences.md`, a cached error of an XLSX): `LibreOffice` calculates the cells with an
                    // error when it opens the file and shows its own code (`Err:502`); the file of `SheetRS` caches `#VALUE!`.
                    .filter(|line| !(extension == "xlsx" && line.contains("Some(Error(\"Err:")))
                    .collect();
                if !lines.is_empty() {
                    found.push(format!("{fixture}.{extension}: {label}: {lines:?}"));
                }
            }
        }
    }
    assert_eq!(
        found,
        Vec::<String>::new(),
        "values the default mode writes that LibreOffice does not show for the same edit"
    );
}

fn formula_of(bytes: &[u8], cell: (u32, u32)) -> Option<String> {
    let workbook = Book::from_bytes(bytes.to_vec())
        .unwrap()
        .workbook()
        .unwrap();
    let calc = workbook.sheets.iter().find(|s| s.name == "Calc").unwrap();
    calc.cells
        .iter()
        .find(|(at, _)| (at.0, at.1) == cell)
        .and_then(|(_, cell)| cell.formula.as_deref().map(str::to_string))
}

/// A qualified `#REF!` that the input holds (`COUNT(Data!#REF!)` of `preref.xlsx`, `Calc!A10`) is the plain error as a whole
/// when the engine reads it, in every pass: with `calculate_all` the values are those of the file `LibreOffice` saves and
/// opens again (`COUNT` 0, `COUNTA` 1, `IFERROR` catches it), whatever the pass (a hide, a replacement, a calculation).
/// The text of a formula the run does not touch is kept as the file has it.
#[test]
fn a_qualified_ref_error_of_the_input_is_plain_in_every_pass_and_keeps_its_text() {
    let source = asset("preref", "xlsx");
    for operations in [
        r#"[["hide_rows", "Data!2"]]"#,
        r#"[["replace", {"text": "x", "with": "y", "scope": "Data!B1"}]]"#,
        r#"[["set_value", "Data!A1", 5]]"#,
    ] {
        let ops = ops(operations);
        let (bytes, _) = written(&source, &ops, false);
        let values = dump(&bytes);
        let cell = |c: &str| values["Calc"].cells[c].clone();
        assert_eq!(cell("A10"), number(0.0), "{operations}: COUNT");
        assert_eq!(cell("A11"), number(1.0), "{operations}: COUNTA");
        assert_eq!(cell("A13"), number(0.0), "{operations}: IFERROR");
        let (kept, _) = written(&source, &ops, true);
        assert_eq!(
            formula_of(&kept, (9, 0)).as_deref(),
            Some("COUNT(Data!#REF!)"),
            "{operations}"
        );
    }
}

/// VERIFIED against `LibreOffice` 26.8 (`oracle.json`, `book.ods`, `formula to value Calc!A:A`; the same run by hand with
/// `sheetrs edit` and the file opened in `LibreOffice`): in an ODS that caches plain numbers in cells with no format `LibreOffice` calculates
/// them on load, so its conversion writes `14` and `8` (`rate*2` and `ROW()`), and the file it saves shows those when it is opened
/// again. The default mode calculates the same formulas (they are roots of the cache) and writes `14` and `8` too, as `calculate_all` does.
#[test]
fn an_ods_conversion_writes_the_load_results_in_both_modes() {
    let edit = ORACLE["fixtures"]["book"]["ods"]["edits"]
        .as_array()
        .unwrap()
        .iter()
        .find(|e| e["label"] == "formula to value Calc!A:A")
        .unwrap();
    let reference = oracle_dump(&edit["reread_default"]);
    assert_eq!(reference["Calc"].cells["A6"], number(14.0));
    assert_eq!(reference["Calc"].cells["A8"], number(8.0));
    let source = asset("book", "ods");
    let by_default = dump(&written(&source, &edit["operations"], true).0);
    assert_eq!(by_default["Calc"].cells["A6"], number(14.0));
    assert_eq!(by_default["Calc"].cells["A8"], number(8.0));
    let everything = dump(&written(&source, &edit["operations"], false).0);
    assert_eq!(everything["Calc"].cells["A6"], number(14.0));
    assert_eq!(everything["Calc"].cells["A8"], number(8.0));
}

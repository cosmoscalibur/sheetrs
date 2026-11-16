//! `SheetRS` against `LibreOffice` on the same file.
//!
//! `sheetrs-io/tests/assets/oracle/*.oracle.json` say what `LibreOffice` itself reports for each cell
//! when it opens that very file (made by `make_oracle.py`): the value, the text it shows, the format
//! code, the formula in its notation. Each file, XLSX or ODS, is compared with its own oracle.
//!
//! The differences the project has approved and documented are listed in `APPROVED` and nowhere else.

use serde_json::{Value, json};
use sheetrs_io::Book;
use sheetrs_protocol::Query;
use sheetrs_query::BookQuery;
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

fn assets() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../sheetrs-io/tests/assets")
}

/// The folder of each fixture, from the root of the repository (the same list as `make_oracle.py`).
fn fixture(file: &str) -> PathBuf {
    let folder = match file.split('.').next().unwrap() {
        "lint_book" => "tests",
        "book" => "sheetrs-edit/tests/assets/structure",
        _ => "sheetrs-io/tests/assets",
    };
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join(folder)
        .join(file)
}

struct Oracle {
    cells: Vec<Value>,
}

fn oracle(file: &str) -> Oracle {
    let path = assets().join("oracle").join(format!("{file}.oracle.json"));
    let json: Value = serde_json::from_slice(&std::fs::read(path).unwrap()).unwrap();
    Oracle {
        cells: json["cells"].as_array().unwrap().clone(),
    }
}

/// The format code `LibreOffice` reports, and the one the readers give, written alike. They are
/// notations of the same format: `LibreOffice` writes its keywords in capitals (`YYYY-MM-DD`), calls
/// the weekday `NNN` and `NN`, quotes or escapes a literal (`"$"` or `\$`), writes the conditions of
/// the sections it infers (`[>0]`), and a file may carry the language mark of a format (`[$-409]`).
fn same_format(libre: &str, ours: Option<&str>) -> bool {
    // `[$EUR-407]` and `[$EUR]` are the currency `EUR`: the language mark only says in which locale
    // it is written.
    let currency = |s: &str| {
        let mut out = String::new();
        let mut rest = s;
        while let Some(at) = rest.find("[$") {
            out.push_str(&rest[..at]);
            let tail = &rest[at..];
            let Some(end) = tail.find(']') else {
                break;
            };
            let inner = &tail[2..end];
            let name = match inner.rfind('-') {
                Some(d) if d > 0 && inner[d + 1..].chars().all(|c| c.is_ascii_hexdigit()) => {
                    &inner[..d]
                }
                _ => inner,
            };
            out.push_str("[$");
            out.push_str(name);
            out.push(']');
            rest = &tail[end + 1..];
        }
        out.push_str(rest);
        out
    };
    let plain = |s: &str| {
        currency(s)
            .replace("[$-409]", "")
            .replace(['\\', '"'], "")
            .replace("[>0]", "")
            .replace("[<0]", "")
            .to_lowercase()
            .replace("nnn", "dddd")
            .replace("nn", "ddd")
    };
    plain(libre) == plain(ours.unwrap_or("General"))
}

fn read_all(file: &str) -> BTreeMap<(String, u32, u32), Value> {
    let query: Query = serde_json::from_value(json!({
        "scope": "workbook",
        "properties": ["value", "formula", "number_format", "display"]
    }))
    .unwrap();
    let source = Book::from_file(fixture(file)).unwrap();
    source
        .query(&query)
        .unwrap()
        .cells
        .into_iter()
        .map(|c| {
            (
                (c.sheet.to_string(), c.row - 1, c.col - 1),
                serde_json::to_value(c).unwrap(),
            )
        })
        .collect()
}

/// The differences from `LibreOffice` that the project approved, each documented where it is
/// described: `(file, sheet, row, column, where it is documented)`. There is none yet.
const APPROVED: &[(&str, &str, u32, u32, &str)] = &[];

/// Differences found that nobody approved and that are not corrected yet: `(file, sheet, row, column, what LibreOffice does)`. Each is listed in the pending decisions; one that stops differing must
/// leave this list (the test says so), and one that is approved moves to `APPROVED`.
const PENDING: &[(&str, &str, u32, u32, &str)] = &[
    (
        "cells.xlsx",
        "T",
        13,
        0,
        "LibreOffice recalculates the cached #VALUE! of an XLSX on load and shows Err:502",
    ),
    (
        "cells.xlsx",
        "T",
        15,
        0,
        "LibreOffice recalculates the cached #VALUE! of an XLSX on load and shows Err:522",
    ),
    (
        "lint_book.xlsx",
        "Loops",
        0,
        0,
        "B1: LibreOffice recalculates the cached #VALUE! of an XLSX on load and shows Err:522",
    ),
    (
        "lint_book.xlsx",
        "Loops",
        1,
        0,
        "B1: LibreOffice recalculates the cached #VALUE! of an XLSX on load and shows Err:522",
    ),
];

/// The built-in number format 14 of an XLSX (`mm-dd-yy` in the standard) is the short date of the
/// locale: `LibreOffice` reads it as `M/D/YYYY` in `en-US` and shows `12/30/1899`, where this reader
/// keeps the literal code of the standard and writes `12-30-99`. It depends on the locale, so it
/// waits for the issue of the locales (#84): `(file, sheet, column)` of the cells that use it.
const LOCALE_PENDING: &[(&str, &str, u32)] =
    &[("dates.xlsx", "Dates", 8), ("dates_1904.xlsx", "Dates", 8)];

fn check(file: &str) -> Vec<String> {
    let ours = read_all(file);
    let mut bad = Vec::new();
    for cell in oracle(file).cells {
        let key = (
            cell["sheet"].as_str().unwrap().to_string(),
            cell["row"].as_u64().unwrap() as u32,
            cell["col"].as_u64().unwrap() as u32,
        );
        let listed = |list: &[(&str, &str, u32, u32, &str)]| {
            list.iter()
                .any(|a| a.0 == file && a.1 == key.0 && (a.2, a.3) == (key.1, key.2))
        };
        let at = format!(
            "{file} {}!{}{}",
            key.0,
            sheetrs_formula::addr::col_name(key.2),
            key.1 + 1
        );
        let mut diffs = Vec::new();
        match ours.get(&key) {
            None => diffs.push(format!("{at}: not read")),
            Some(mine) => {
                // What depends on the clock is not comparable: `LibreOffice` recalculates it when
                // it opens the file.
                let volatile = cell["formula"].as_str().is_some_and(|f| {
                    ["NOW()", "TODAY()", "RAND()", "RANDBETWEEN("]
                        .iter()
                        .any(|v| f.to_uppercase().contains(v))
                });
                let libre_shows = cell["string"].as_str().unwrap();
                let shows = mine["display"].as_str().unwrap_or_default();
                if libre_shows != shows && !volatile {
                    diffs.push(format!(
                        "{at}: shows: LibreOffice {libre_shows:?}, SheetRS {shows:?}"
                    ));
                }
                let libre_format = cell["format"].as_str().unwrap();
                if !same_format(libre_format, mine["number_format"].as_str()) {
                    diffs.push(format!(
                        "{at}: format: LibreOffice {libre_format:?}, SheetRS {:?}",
                        mine["number_format"].as_str().unwrap_or("General")
                    ));
                }
                // The cells of an array formula are checked in `libreoffice_arrays.rs`.
                if let Some(libre) = cell["formula"].as_str().filter(|f| !f.starts_with('{')) {
                    let sheet = key.0.as_str();
                    let format = if std::path::Path::new(file)
                        .extension()
                        .is_some_and(|e| e.eq_ignore_ascii_case("ods"))
                    {
                        sheetrs_io::reader::Format::Ods
                    } else {
                        sheetrs_io::reader::Format::Xlsx
                    };
                    let text = sheetrs_query::text::TextContext::new(
                        None,
                        sheetrs_formula::DateSystem::default(),
                    )
                    .unwrap()
                    .for_format(format)
                    .for_book(&Book::from_file(fixture(file)).unwrap());
                    // The folder of the document, where `LibreOffice` resolves external files.
                    let folder = std::fs::canonicalize(fixture(file).parent().unwrap()).unwrap();
                    let libre = libre.replace("<folder>", &folder.display().to_string());
                    let libre = libre.as_str();
                    let ours =
                        text.formula_text(mine["formula"].as_str().unwrap_or_default(), sheet);
                    if libre != ours {
                        diffs.push(format!(
                            "{at}: formula: LibreOffice {libre:?}, SheetRS {ours:?}"
                        ));
                    }
                }
                if cell["error"].as_u64() == Some(0)
                    && cell["type"] != "TEXT"
                    && cell["result"].as_u64() != Some(2)
                    && !volatile
                    && let (Some(l), Some(m)) = (cell["value"].as_f64(), mine["value"].as_f64())
                    && (l - m).abs() > 1e-9 * l.abs().max(1.0)
                {
                    diffs.push(format!("{at}: value: LibreOffice {l}, SheetRS {m}"));
                }
            }
        }
        if listed(APPROVED) {
            continue;
        }
        let locale_listed = LOCALE_PENDING
            .iter()
            .any(|a| a.0 == file && a.1 == key.0 && a.2 == key.2);
        if listed(PENDING) || locale_listed {
            if diffs.is_empty() {
                bad.push(format!(
                    "{at}: matches LibreOffice now: take it out of PENDING"
                ));
            }
        } else {
            bad.extend(diffs);
        }
    }
    bad
}

/// What a `LibreOffice` with the `es-CO` regional setting shows for `locale_es-CO.{xlsx,ods}` (made
/// by `make_locale.py`, recorded by `make_oracle.py`): the text of every cell, in the locale of the
/// workbook: the one given to its `Book` (neither format carries the regional setting of the
/// machine that saved it, and an ODS's default cell style does not change what `LibreOffice` shows).
/// The format codes the oracle reports are written in
/// Spanish there (`DD/MM/AAAA`), so only what the cells show is compared.
fn display_in_the_locale_of(file: &str, locale: Option<&str>) -> Vec<String> {
    let query: Query = serde_json::from_value(json!({
        "scope": "workbook",
        "properties": ["display"]
    }))
    .unwrap();
    let mut book = Book::from_file(fixture(file)).unwrap();
    if let Some(tag) = locale {
        book = book.with_locale(tag.parse().unwrap());
    }
    let shown: BTreeMap<(String, u32, u32), String> = book
        .query(&query)
        .unwrap()
        .cells
        .into_iter()
        .map(|c| {
            (
                (c.sheet.to_string(), c.row - 1, c.col - 1),
                c.display.unwrap_or_default(),
            )
        })
        .collect();
    let mut bad = Vec::new();
    for cell in oracle(file).cells {
        let key = (
            cell["sheet"].as_str().unwrap().to_string(),
            cell["row"].as_u64().unwrap() as u32,
            cell["col"].as_u64().unwrap() as u32,
        );
        let libre = cell["string"].as_str().unwrap();
        let ours = shown.get(&key).map_or("", String::as_str);
        if libre != ours {
            bad.push(format!(
                "{file} {}!{}{}: LibreOffice shows {libre:?}, SheetRS {ours:?}",
                key.0,
                sheetrs_formula::addr::col_name(key.2),
                key.1 + 1
            ));
        }
    }
    bad
}

#[test]
fn the_display_of_a_book_follows_its_locale() {
    let mut bad = display_in_the_locale_of("locale_es-CO.xlsx", Some("es-CO"));
    bad.extend(display_in_the_locale_of("locale_es-CO.ods", Some("es-CO")));
    assert!(
        bad.is_empty(),
        "{} differences:\n{}",
        bad.len(),
        bad.join("\n")
    );
}

#[test]
fn the_locale_given_to_the_book_changes_the_display_of_a_query() {
    let spanish = |file: &str, locale: Option<&str>| {
        let mut book = Book::from_file(fixture(file)).unwrap();
        if let Some(tag) = locale {
            book = book.with_locale(tag.parse().unwrap());
        }
        let query: Query = serde_json::from_value(json!({
            "scope": {"range": "Numbers!C1"},
            "properties": ["display"]
        }))
        .unwrap();
        book.query(&query).unwrap().cells[0]
            .display
            .clone()
            .unwrap()
    };
    assert_eq!(spanish("locale_es-CO.ods", Some("es-CO")), "1.234,50");
    assert_eq!(spanish("locale_es-CO.ods", Some("en-US")), "1,234.50");
    // The locale of a single query wins over the one of the book.
    let query: Query = serde_json::from_value(json!({
        "scope": {"range": "Numbers!C1"},
        "properties": ["display"],
        "locale": "en-US"
    }))
    .unwrap();
    let book = Book::from_file(fixture("locale_es-CO.ods"))
        .unwrap()
        .with_locale("es-CO".parse().unwrap());
    assert_eq!(
        book.query(&query).unwrap().cells[0].display.as_deref(),
        Some("1,234.50")
    );
}

#[test]
fn what_each_file_shows_is_what_libreoffice_shows() {
    let mut bad = Vec::new();
    for file in [
        "cells.xlsx",
        "cells.ods",
        "inherit.xlsx",
        "inherit.ods",
        "inherit_more.xlsx",
        "inherit_more.ods",
        "inherit_text.xlsx",
        "inherit_text.ods",
        "formulas.xlsx",
        "formulas.ods",
        "formats.xlsx",
        "formats.ods",
        "book.xlsx",
        "book.ods",
        "errors.xlsx",
        "errors.ods",
        "lint_book.xlsx",
        "lint_book.ods",
        "dates.xlsx",
        "dates.ods",
        "dates_1904.xlsx",
        "dates_1904.ods",
        "dates_1900.ods",
    ] {
        bad.extend(check(file));
    }
    assert!(
        bad.is_empty(),
        "{} differences:\n{}",
        bad.len(),
        bad.join("\n")
    );
}

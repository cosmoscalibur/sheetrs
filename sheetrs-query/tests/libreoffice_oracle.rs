//! `SheetRS` against `LibreOffice` on the same file.
//!
//! `sheetrs-io/tests/assets/oracle/*.oracle.json` say what `LibreOffice` itself reports for each cell
//! when it opens that very file (made by `make_oracle.py`): the value, the text it shows, the format
//! code, the formula in its notation. Each file, XLSX or ODS, is compared with its own oracle; the
//! other format of the workbook is not a reference.
//!
//! The differences the project has approved and documented are listed in `APPROVED` and nowhere else.

use serde_json::{Value, json};
use sheetrs_io::reader::Source;
use sheetrs_protocol::Query;
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

fn assets() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../sheetrs-io/tests/assets")
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
    let plain = |s: &str| {
        s.replace("[$-409]", "")
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
    let source = Source::from_file(assets().join(file)).unwrap();
    sheetrs_query::run(&source, &query)
        .unwrap()
        .cells
        .into_iter()
        .map(|c| {
            (
                (c.sheet.clone(), c.row - 1, c.col - 1),
                serde_json::to_value(c).unwrap(),
            )
        })
        .collect()
}

/// A formula in the notation of `LibreOffice`, without the mark of a sheet that is absolute
/// (`$Sheet.A1` is `Sheet.A1` here). `LibreOffice` tells a sheet that is absolute from one that is
/// relative (an ODS keeps both: `[$T.A1]` and `[T.A1]`; an XLSX has only the first), and the readers
/// give a formula in Excel's notation, which has no relative sheets. So a search of an ODS formula
/// sees `$T.A1` where `LibreOffice` shows `T.A1`: that difference is pending (see the pending
/// decisions) and the comparison here leaves the mark out.
fn without_sheet_mark(formula: &str) -> String {
    let marked = regex::Regex::new(r"\$('(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_]*)\.").unwrap();
    marked.replace_all(formula, "${1}.").into_owned()
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
        "cells.xlsx",
        "T",
        19,
        0,
        "a formula whose result is a reference takes the format of the referenced cell (INDIRECT)",
    ),
    (
        "formulas.xlsx",
        "F",
        17,
        0,
        "a range in a cell holds #VALUE! in the file; the reader takes it for the marker of a dynamic array (#62) and gives nothing",
    ),
    ("formulas.xlsx", "F", 21, 0, "same as F18 (#62)"),
    ("formulas.xlsx", "F", 23, 0, "same as F18 (#62)"),
    ("formulas.xlsx", "F", 40, 0, "same as F18 (#62)"),
    (
        "cells.xlsx",
        "T",
        11,
        0,
        "an unknown function is shown as it was written (`foo()`), the readers write it in capitals",
    ),
    ("cells.ods", "T", 11, 0, "same as cells.xlsx T12"),
    (
        "cells.xlsx",
        "T",
        17,
        0,
        "a reference to the own sheet that was written with the sheet keeps it (`$T.A1`); the notation here leaves the own sheet out",
    ),
    (
        "cells.ods",
        "T",
        17,
        0,
        "same as cells.xlsx T18, with a sheet that is relative (`T.A1`)",
    ),
    (
        "formulas.ods",
        "F",
        6,
        0,
        "an ODS range with the sheet on both ends (`$T.A1:$T.B3`) is written with it once (`$T.A1:B3`)",
    ),
    (
        "formulas.ods",
        "F",
        7,
        0,
        "same as F7, with a relative sheet",
    ),
    (
        "formulas.ods",
        "F",
        17,
        0,
        "same as F7, with a quoted sheet name",
    ),
];

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
                let libre_shows = cell["string"].as_str().unwrap();
                let shows = mine["display"].as_str().unwrap_or_default();
                if libre_shows != shows {
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
                if let Some(libre) = cell["formula"].as_str() {
                    let sheet = key.0.as_str();
                    let text = sheetrs_query::text::TextContext::new(None, false).unwrap();
                    let ours =
                        text.formula_text(mine["formula"].as_str().unwrap_or_default(), sheet);
                    if without_sheet_mark(libre) != without_sheet_mark(&ours) {
                        diffs.push(format!(
                            "{at}: formula: LibreOffice {libre:?}, SheetRS {ours:?}"
                        ));
                    }
                }
                if cell["error"].as_u64() == Some(0)
                    && cell["type"] != "TEXT"
                    && cell["result"].as_u64() != Some(2)
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
        if listed(PENDING) {
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

#[test]
fn what_each_file_shows_is_what_libreoffice_shows() {
    let mut bad = Vec::new();
    for file in [
        "cells.xlsx",
        "cells.ods",
        "formulas.xlsx",
        "formulas.ods",
        "formats.xlsx",
        "formats.ods",
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

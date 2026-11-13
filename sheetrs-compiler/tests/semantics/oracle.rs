//! Cross-check against `LibreOffice`: `tests/assets/semantics.{xlsx,ods}` hold
//! ~200 formulas whose cached values were computed by `LibreOffice` (see
//! `scripts/gen_semantics_fixture.py`). The engine must reproduce them from the
//! formulas alone, from each file on its own: the oracle of a file is what `LibreOffice` gives when it
//! opens that file. An XLSX caches `#VALUE!` for every error that has a code of its own, so the
//! generator also asks `LibreOffice` to recalculate the XLSX itself and keeps what differs from the
//! cache in `<name>.xlsx.recalc.json` (`recalculated_oracle` in the script).

use serde_json::Value as Json;
use sheetrs_compiler::{BindingSpec, Spec, compile, loader};
use sheetrs_engine::Overrides;
use sheetrs_engine::addr::CellRef;
use sheetrs_engine::value::Value;
use std::path::Path;

/// What `LibreOffice` gives for the rows of an XLSX when it opens and recalculates that file, where
/// that is not what the file caches (`{row: {"error": "Err:502"} | {"number": n} | {"bool": b} |
/// {"text": t}}`, rows 1-based): the codes of the errors, and the rows whose formula the XLSX
/// export rewrites (`T1:T1` as `T1`, `IF({1;0})` as `IF({1;0},TRUE())`), judged as the file writes
/// them.
fn recalculated(file: &str) -> std::collections::HashMap<u32, Json> {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(format!("{file}.recalc.json"));
    let text = std::fs::read_to_string(&path).unwrap_or_else(|e| {
        panic!(
            "{}: {e} (scripts/gen_semantics_fixture.py makes it)",
            path.display()
        )
    });
    let json: Json = serde_json::from_str(&text).unwrap();
    json["rows"]
        .as_object()
        .unwrap()
        .iter()
        .map(|(row, entry)| (row.parse().unwrap(), entry.clone()))
        .collect()
}

/// The error of `got` against the one `LibreOffice` shows (`Err:502`, `#N/A`…): the Excel name and,
/// when `LibreOffice` has a code of its own, that code.
fn same_error(theirs: &str, got: &Json) -> Result<(), String> {
    use sheetrs_io::reader::ExcelError;
    use std::str::FromStr;
    let Json::Object(o) = got else {
        return Err(format!("engine: {got}"));
    };
    let kind = ExcelError::from_str(theirs).map_err(|e| e.to_string())?;
    let (name, code) = match kind {
        ExcelError::Libre(522) => ("#CIRC!".to_string(), None),
        ExcelError::Libre(c) => ("#VALUE!".to_string(), Some(format!("Err:{c}"))),
        other => (other.to_string(), None),
    };
    let ours_name = o.get("error").and_then(Json::as_str).unwrap_or("");
    let ours_code = o.get("code").and_then(Json::as_str).map(String::from);
    if ours_name == name && ours_code == code {
        Ok(())
    } else {
        Err(format!(
            "engine: {ours_name} {ours_code:?}, LibreOffice: {name} {code:?}"
        ))
    }
}

/// `locale`: the regional settings `LibreOffice` ran with when it computed `file`.
fn check(file: &str, locale: &str) {
    check_rows(file, locale, 150, &[]);
}

/// `known`: labels of rows where the engine is known to differ from `LibreOffice` in this file. They
/// must still differ: a row that agrees is an error until it is taken off the list.
fn check_rows(file: &str, locale: &str, min_rows: usize, known: &[&str]) {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(file);
    let model = loader::load(&sheetrs_io::Book::from_file(&path).unwrap()).unwrap();
    let calc = model.sheet_id("Calc").unwrap();
    let mut rows: Vec<u32> = model
        .cells
        .iter()
        .filter(|(at, c)| at.sheet == calc && at.col == 0 && c.formula.is_some())
        .map(|(at, _)| at.row)
        .collect();
    rows.sort_unstable();
    assert!(
        rows.len() >= min_rows,
        "fixture looks truncated: {}",
        rows.len()
    );
    let last = *rows.last().unwrap();
    let label_of = |row: u32| {
        model
            .cells
            .get(&CellRef::new(calc, row, 1))
            .map_or(String::new(), |c| match &c.value {
                Value::Text(t) => t.to_string(),
                _ => String::new(),
            })
    };
    // A date written without a year is of the year the oracle ran in: its `TODAY()`.
    let now = rows.iter().find_map(|&r| {
        match (
            &model.cells[&CellRef::new(calc, r, 0)].value,
            label_of(r).as_str(),
        ) {
            (Value::Number(n), "TODAY()") => Some(*n),
            _ => None,
        }
    });
    let spec = Spec {
        outputs: vec![BindingSpec::parse(&format!("Calc!A1:A{}", last + 1))],
        fold: false,
        now,
        locale: Some(
            sheetrs_compiler::locale::from_tag(
                locale,
                &sheetrs_compiler::locale::LocaleOverride::default(),
            )
            .unwrap()
            .0,
        ),
        ..Spec::default()
    };
    let (program, report) = compile(&model, &spec).unwrap();
    let _ = report;
    let out = program
        .calculate(&Json::Null, &Overrides::default())
        .unwrap();
    let col = out[&format!("Calc!A1:A{}", last + 1)].as_array().unwrap();

    let judged = if file.ends_with("xlsx") {
        recalculated(file)
    } else {
        std::collections::HashMap::default()
    };
    let mut bad = Vec::new();
    let mut differ = std::collections::HashSet::new();
    for &row in &rows {
        let cell = &model.cells[&CellRef::new(calc, row, 0)];
        let label = label_of(row);
        let got = &col[row as usize][0];
        // An XLSX is judged on what `LibreOffice` gives when it opens that file and recalculates
        // it (`recalculated`): the code of an error, and the rows its export rewrote.
        if let Some(entry) = judged.get(&(row + 1)) {
            let verdict = if let Some(theirs) = entry.get("error").and_then(Json::as_str) {
                same_error(theirs, got)
            } else {
                let value = if let Some(n) = entry.get("number").and_then(Json::as_f64) {
                    Value::Number(n)
                } else if let Some(b) = entry.get("bool").and_then(Json::as_bool) {
                    Value::Bool(b)
                } else {
                    Value::Text(entry["text"].as_str().unwrap_or("").into())
                };
                if crate::common::matches_oracle(&value, got) {
                    Ok(())
                } else {
                    Err(format!("oracle: {value:?}\n        engine: {got}"))
                }
            };
            if let Err(why) = verdict {
                if known.contains(&label.as_str()) {
                    differ.insert(label);
                } else {
                    bad.push(format!("  A{:<4} {label}\n        {why}", row + 1));
                }
            }
            continue;
        }
        // An ODS says which error LibreOffice gave (`Err:502`, `Err:523`; `Err:519` is `#VALUE!`):
        // the engine must give that very code, in every row.
        if file.ends_with("ods")
            && let (Value::Error(a), Json::Object(o)) = (&cell.value, got)
        {
            let theirs = a.libre_code();
            let ours = o
                .get("code")
                .and_then(Json::as_str)
                .map(String::from)
                .or_else(|| {
                    (o.get("error").and_then(Json::as_str) == Some("#CIRC!"))
                        .then(|| "Err:522".to_string())
                });
            if theirs != ours {
                if known.contains(&label.as_str()) {
                    differ.insert(label);
                } else {
                    bad.push(format!(
                        "  A{:<4} {label}\n        LibreOffice: {theirs:?} ({a:?})\n        engine: {ours:?}",
                        row + 1
                    ));
                }
                continue;
            }
            // The code is the one of `LibreOffice`: the error is the same, whatever name the loader
            // gives its kind .
            continue;
        }
        if !crate::common::matches_oracle(&cell.value, got) {
            if known.contains(&label.as_str()) {
                differ.insert(label);
            } else {
                bad.push(format!(
                    "  A{:<4} {label}\n        oracle: {:?}\n        engine: {got}",
                    row + 1,
                    cell.value
                ));
            }
        }
    }
    for k in known {
        assert!(
            differ.contains(*k),
            "{file}: `{k}` agrees with LibreOffice now: take it off the list of known differences"
        );
    }
    assert!(
        bad.is_empty(),
        "{file}: {} mismatches\n{}",
        bad.len(),
        bad.join("\n")
    );
}

/// `SUM(Data!A1:Wide!A1048576)` is written by `LibreOffice` as `SUM(Data:Wide!A:A)`; the headless
/// `LibreOffice` that recalculates the file for this oracle reads the whole-column form over two
/// sheets as `#NAME?` (a file written by `LibreOffice` holds the form, and Excel writes it: the
/// engine reads it, FIDELITY.md).
#[test]
fn the_semantics_xlsx_calculates_every_formula_as_libreoffice_does() {
    check_rows("semantics.xlsx", "en", 150, &["SUM(Data!A1:Wide!A1048576)"]);
}

#[test]
fn the_semantics_ods_calculates_every_formula_as_libreoffice_does() {
    check("semantics.ods", "en");
}

/// Locales with a fixture of their own, computed by a `LibreOffice` running with that regional
/// setting (`scripts/gen_semantics_fixture.py`, `LOCALE_FIXTURES`).
fn check_locale(tag: &str) {
    check_rows(&format!("semantics_{tag}.xlsx"), tag, 150, &[]);
    check_rows(&format!("semantics_{tag}.ods"), tag, 150, &[]);
}

#[test]
fn spanish_of_spain_matches_libreoffice() {
    check_locale("es-ES");
}

#[test]
fn spanish_of_colombia_matches_libreoffice() {
    check_locale("es-CO");
}

#[test]
fn spanish_of_mexico_matches_libreoffice() {
    check_locale("es-MX");
}

#[test]
fn british_english_matches_libreoffice() {
    check_locale("en-GB");
}

/// A spreadsheet created in `LibreOffice` (ODF default settings, `LibreOffice`'s own function
/// names such as the native `CEILING`), not imported from XLSX.
///
/// The formulas are the ones of `semantics.xlsx` typed again, and `LibreOffice` reads some of them
/// otherwise than the import did; the engine answers what the typed text says in this row:
/// - `SUM(Data!A1:Wide!A1048576)`: typed, the whole-column form over two sheets is `#NAME?` in the
///   headless `LibreOffice` (a file written by `LibreOffice` holds the form the engine reads).
#[test]
fn native_ods_matches_libreoffice() {
    check_rows("native.ods", "en", 150, &["SUM(Data!A1:Wide!A1048576)"]);
}

/// Documents created in `LibreOffice` with the null dates that are not its default (the date of
/// serial 0): 1900-01-01 (`StarCalc` 1.0) and 1904-01-01. The date system is the document's, one of the
/// three `LibreOffice` Calc offers.
#[test]
fn null_dates_match_libreoffice() {
    for file in ["null_date_1900.ods", "null_date_1904.ods"] {
        check_rows(file, "en", 20, &[]);
    }
}

/// Complex numbers written as text are compared by precision, not by characters.
#[test]
fn complex_text_is_compared_by_precision() {
    let matches = |oracle: &str, engine: &str| {
        crate::common::matches_oracle(
            &Value::Text(oracle.into()),
            &Json::String(engine.to_string()),
        )
    };
    // The last digit of a 15-digit part, as a different C library rounds it.
    assert!(matches("27.2899171971277i", "27.2899171971278i"));
    assert!(matches(
        "0.0202237833442819+0.274925301029484i",
        "0.020223783344282+0.274925301029484i"
    ));
    assert!(matches("1-i", "1-1i"));
    // A part that differs, another suffix or a text that is not a number do not match.
    assert!(!matches("27.2899171971277i", "27.289917197i"));
    assert!(!matches("1+2i", "1+2j"));
    assert!(!matches("abc", "abd"));
}

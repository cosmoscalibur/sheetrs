//! Cross-check against LibreOffice: `tests/assets/semantics.{xlsx,ods}` hold
//! ~200 formulas whose cached values were computed by LibreOffice (see
//! `scripts/gen_semantics_fixture.py`). The engine must reproduce them from the
//! formulas alone, from both container formats.

use serde_json::Value as Json;
use sheetrs_compiler::addr::CellRef;
use sheetrs_compiler::value::Value;
use sheetrs_compiler::{BindingSpec, Overrides, Spec, compile, loader};
use std::path::Path;

/// Formulas where the oracle itself cannot be compared as is (the reference is
/// LibreOffice for both files): (label, file suffix
/// or "" for all, reason).
const ORACLE_GAPS: &[(&str, &str, &str)] = &[
    (
        "_xlfn.F.INV(1,5,8)",
        "",
        "LibreOffice answers the end of its search range (2.37E+81); the engine does not reproduce it (`notes::F_INV_ONE`)",
    ),
    (
        "Data!A1:A3*2",
        "xlsx",
        "sheetrs reads a cached #VALUE! of a formula holding `:` as no value (#62)",
    ),
    ("SUM(Data!A1:A3*2)", "xlsx", "same"),
    ("SUM(Data!A1:A3>3)", "xlsx", "same"),
    (
        "\"[\"&100000000000000000000&\"]\"",
        "xlsx",
        "LibreOffice writes the exponent with three digits (1E+020)",
    ),
];

/// LibreOffice's XLSX export writes every error of an invalid argument as
/// `#VALUE!` (and sheetrs reads a cached `#VALUE!` of a formula with `:` as no
/// value, #62), so the kind an Excel-semantics engine must give cannot be read
/// from it. The ODS twin keeps the real code: when LibreOffice says "error" there,
/// an engine error of any kind is accepted for the XLSX row.
fn ods_errors(twin: &str) -> std::collections::HashSet<String> {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(twin);
    let model = loader::load(&path).unwrap();
    let calc = model.sheet_id("Calc").unwrap();
    model
        .cells
        .iter()
        .filter(|(at, c)| at.sheet == calc && at.col == 0 && matches!(c.value, Value::Error(_)))
        .filter_map(
            |(at, _)| match &model.cells.get(&CellRef::new(calc, at.row, 1))?.value {
                Value::Text(t) => Some(t.to_string()),
                _ => None,
            },
        )
        .collect()
}

/// `locale`: the regional settings LibreOffice ran with when it computed `file`.
fn check(file: &str, locale: &str) {
    check_rows(file, locale, 150);
}

fn check_rows(file: &str, locale: &str, min_rows: usize) {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(file);
    let model = loader::load(&path).unwrap();
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
            sheetrs_compiler::locale::from_tag(locale, &Default::default())
                .unwrap()
                .0,
        ),
        ..Spec::default()
    };
    let (program, report) = compile(&model, &spec).unwrap();
    let _ = report;
    let out = program.execute(&Json::Null, &Overrides::default()).unwrap();
    let col = out[&format!("Calc!A1:A{}", last + 1)].as_array().unwrap();

    let twin_errors = if file.ends_with("xlsx") {
        ods_errors(&file.replace(".xlsx", ".ods"))
    } else {
        Default::default()
    };
    let mut bad = Vec::new();
    for &row in &rows {
        let cell = &model.cells[&CellRef::new(calc, row, 0)];
        let label =
            model
                .cells
                .get(&CellRef::new(calc, row, 1))
                .map_or(String::new(), |c| match &c.value {
                    Value::Text(t) => t.to_string(),
                    _ => String::new(),
                });
        let got = &col[row as usize][0];
        if let Some((_, _, _)) = ORACLE_GAPS
            .iter()
            .find(|(f, ext, _)| *f == label && (ext.is_empty() || file.ends_with(ext)))
        {
            continue;
        }
        let engine_error = matches!(got, Json::Object(_));
        let oracle_unreadable = matches!(cell.value, Value::Error(_) | Value::Empty);
        if engine_error && oracle_unreadable && twin_errors.contains(&label) {
            continue;
        }
        if !crate::common::matches_oracle(&cell.value, got) {
            bad.push(format!(
                "  A{:<4} {label}\n        oracle: {:?}\n        engine: {got}",
                row + 1,
                cell.value
            ));
        }
    }
    assert!(
        bad.is_empty(),
        "{file}: {} mismatches\n{}",
        bad.len(),
        bad.join("\n")
    );
}

#[test]
fn xlsx_matches_libreoffice() {
    check("semantics.xlsx", "en");
}

#[test]
fn ods_matches_libreoffice() {
    check("semantics.ods", "en");
}

/// Locales with a fixture of their own, computed by a LibreOffice running with that regional
/// setting (`scripts/gen_semantics_fixture.py`, `LOCALE_FIXTURES`).
fn check_locale(tag: &str) {
    check_rows(&format!("semantics_{tag}.xlsx"), tag, 150);
    check_rows(&format!("semantics_{tag}.ods"), tag, 150);
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

/// A spreadsheet created in LibreOffice (ODF default settings, LibreOffice's own function
/// names such as the native `CEILING`), not imported from XLSX.
#[test]
fn native_ods_matches_libreoffice() {
    check("native.ods", "en");
}

/// Documents created in LibreOffice with another null date (the date of serial 0): the date
/// system is the document's, whatever it is.
#[test]
fn null_dates_match_libreoffice() {
    for file in [
        "null_date_1900.ods",
        "null_date_1904.ods",
        "null_date_1950.ods",
    ] {
        check_rows(file, "en", 20);
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

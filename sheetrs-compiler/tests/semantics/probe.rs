//! Fidelity probes: formulas on which Excel and LibreOffice are known or suspected to differ
//! (`scripts/probe_behavior.py`, documented in docs/FIDELITY.md). The ODS files hold
//! LibreOffice's answers, which the engine must reproduce; the cases where it differs from
//! Excel are pinned below.

use serde_json::{Value as Json, json};
use sheetrs_compiler::addr::CellRef;
use sheetrs_compiler::options::SearchMode;
use sheetrs_compiler::value::Value;
use sheetrs_compiler::{BindingSpec, Model, Overrides, Spec, compile, loader};
use std::path::Path;

/// Probes the engine deliberately does not reproduce, with the reason.
const NOT_REPRODUCED: &[(&str, &str)] = &[
    (
        "SEARCH(\"a*\",\"xabc\")",
        "a regular expression that can match the empty text: outside the subset, the build stops",
    ),
    (
        "COUNTIF(Data!B1:B8,\"*an*\")",
        "an invalid regular expression: outside the subset, the build stops",
    ),
    (
        "COUNTIF(Data!B1:B8,\"a*\")",
        "a regular expression that can match the empty text: outside the subset, the build stops",
    ),
];

/// Error-valued probes where LibreOffice's error kind is not meaningful
/// (every invalid argument is Err:502): only "is an error" is compared.
const ANY_ERROR: &[&str] = &[
    "CEILING(2.1,-2)",
    "FLOOR(2.1,-2)",
    "LOG(1,1)",
    "FACT(171)",
    "SQRT(-1)",
];

fn run(model: &Model, n: u32) -> Vec<Json> {
    let spec = Spec {
        outputs: vec![BindingSpec::parse(&format!("Calc!A1:A{n}"))],
        fold: false,
        // The probes that are outside what the engine reproduces are listed below; the build would
        // stop at them.
        inspect: true,
        ..Spec::default()
    };
    let (p, _) = compile(model, &spec).unwrap();
    let out = p.execute(&Json::Null, &Overrides::default()).unwrap();
    out[&p.outputs[0].name]
        .as_array()
        .unwrap()
        .iter()
        .map(|r| r[0].clone())
        .collect()
}

fn label(m: &Model, calc: u16, row: u32) -> String {
    match &m.cells[&CellRef::new(calc, row, 1)].value {
        Value::Text(t) => t.to_string(),
        _ => String::new(),
    }
}

fn reproduces(file: &str, expect: impl Fn(&Model)) {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(file);
    let m = loader::load(&path).unwrap();
    expect(&m);
    let calc = m.sheet_id("Calc").unwrap();
    let n = m
        .cells
        .keys()
        .filter(|a| a.sheet == calc && a.col == 0)
        .count() as u32;
    let got = run(&m, n);
    let mut bad = Vec::new();
    for row in 0..n {
        let l = label(&m, calc, row);
        if NOT_REPRODUCED.iter().any(|(f, _)| *f == l) {
            continue;
        }
        let expected = &m.cells[&CellRef::new(calc, row, 0)].value;
        let g = &got[row as usize];
        let ok = if ANY_ERROR.contains(&l.as_str()) {
            matches!(expected, Value::Error(_)) && g.get("error").is_some()
        } else {
            crate::common::matches_oracle(expected, g)
        };
        if !ok {
            bad.push(format!("  {l:<46} LibreOffice: {expected:?}   engine: {g}"));
        }
    }
    assert!(
        bad.is_empty(),
        "{file}: {} mismatches:\n{}",
        bad.len(),
        bad.join("\n")
    );
}

#[test]
fn the_engine_reproduces_libreoffice_on_the_probes() {
    reproduces("probe.ods", |m| {
        let s = m.settings.search.unwrap();
        assert_eq!(
            (s.mode, s.case_sensitive, s.whole_cell),
            (SearchMode::Wildcards, false, true)
        );
    });
}

/// Native ODS settings the file states: regular expressions, case-sensitive,
/// partial-cell matching.
#[test]
fn document_search_settings_regex_mode() {
    reproduces("probe_regex.ods", |m| {
        let s = m.settings.search.unwrap();
        assert_eq!(
            (s.mode, s.case_sensitive, s.whole_cell),
            (SearchMode::Regex, true, false)
        );
    });
}

#[test]
fn document_search_settings_literal_mode() {
    reproduces("probe_literal.ods", |m| {
        let s = m.settings.search.unwrap();
        assert_eq!(
            (s.mode, s.case_sensitive, s.whole_cell),
            (SearchMode::Literal, false, true)
        );
    });
}

/// The behaviors that differ from Excel's, pinned: the answers are LibreOffice's.
#[test]
fn libreoffice_behaviors_that_differ_from_excel_are_pinned() {
    let cases: &[(&str, Json)] = &[
        ("0^0", json!(1)),
        ("(-8)^(1/3)", json!(-2)),
        ("TRUE&\"\"", json!("1")),
        ("TEXT(TRUE,\"0\")", json!("1")),
        ("\"a\"<TRUE", json!(false)),
        ("TRUE=1", json!(true)),
        ("FALSE<1", json!(true)),
        ("ISNUMBER(TRUE)", json!(true)),
        ("SUM(\"1\")", crate::common::err("#VALUE!")),
        ("\"[\"&1E-10&\"]\"", json!("[0.0000000001]")),
        ("\"[\"&1E+20&\"]\"", json!("[1E+020]")),
        ("DATE(1900,1,1)", json!(2)),
        ("DATE(1900,2,28)", json!(60)),
        ("YEAR(0)", json!(1899)),
        ("DATE(1850,1,1)", json!(-18260)),
        ("DATE(29,1,1)", json!(47119)),
        ("DATEVALUE(\"1 Mar 2024\")", crate::common::err("#VALUE!")),
        ("DATEVALUE(\"2024/03/01\")", crate::common::err("#VALUE!")),
        ("DATEVALUE(\"1-Mar-2024\")", json!(45352)),
        ("UPPER(\"ß\")", json!("ẞ")),
        ("UPPER(\"straße\")", json!("STRAẞE")),
    ];
    for (f, lo) in cases {
        assert_eq!(&crate::common::calc(f), lo, "{f}");
    }
}

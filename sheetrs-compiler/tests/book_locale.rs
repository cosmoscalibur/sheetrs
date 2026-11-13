//! The locale of a calculation is a property of the workbook (`Book::with_locale`; the default of its
//! format when it has none): the text a number becomes (`TEXT`, the conversion in a `&`) is the one
//! `LibreOffice` writes under that locale. The oracle is `semantics_es-CO.{xlsx,ods}`: made by a
//! `LibreOffice` running with the `es-CO` regional setting, with the cases of issues #84 and #106
//! among its rows, that the engine must give again from the formulas alone.

use serde_json::{Value as Json, json};
use sheetrs_compiler::memory::{BookCompile, Options};
use sheetrs_engine::ir::Mode;
use sheetrs_engine::value::Value;
use sheetrs_io::Book;
use std::path::{Path, PathBuf};

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(name)
}

/// The rows of the `Calc` sheet of the oracle by the text of their formula (column B).
fn row_of(file: &str, formula: &str) -> u32 {
    let book = Book::from_file(asset(file)).unwrap();
    let model = sheetrs_compiler::loader::load(&book).unwrap();
    let calc = model.sheet_id("Calc").unwrap();
    model
        .cells
        .iter()
        .find_map(|(at, cell)| {
            (at.sheet == calc
                && at.col == 1
                && matches!(&cell.value, Value::Text(t) if &**t == formula))
            .then_some(at.row + 1)
        })
        .unwrap_or_else(|| panic!("{file} has no row for {formula}"))
}

/// What `book` calculates for the formula that the oracle `file` has in a row.
fn calculated(file: &str, book: &Book, options: &Options, formula: &str) -> Json {
    let range = format!("Calc!A{}", row_of(file, formula));
    let request = json!({ "ranges": [range] });
    let done = book
        .calculate(&request, options, Mode::Values, false)
        .unwrap();
    done.answer.unwrap().to_json()["valueRanges"][0]["values"][0][0].clone()
}

/// The four cases of issues #84 and #106, as `LibreOffice` shows them under `es-CO` and under `en-US`.
const CASES: [(&str, &str, &str); 4] = [
    (
        r#"TEXT(350*49799,"[$$-en-US]#.###.##0")"#,
        "$17.429.650",
        "$17429650.000.000",
    ),
    (
        r#""Componente inflacionario "&(0.5543*100)&"%""#,
        "Componente inflacionario 55,43%",
        "Componente inflacionario 55.43%",
    ),
    (
        r#"TEXT(192*49799,"[$$-en-US]#.###.##0")&" ("&TEXT(192*49799/12,"[$$-en-US]#.##0")&")""#,
        "$9.561.408 ($796.784)",
        "$9561408.000.000 ($796784.000)",
    ),
    (
        r#"TEXT(0,"[$$-en-US]#,##0;([$$-en-US]#,##0)")"#,
        "$,000",
        "$0",
    ),
];

fn book(file: &str, locale: Option<&str>) -> Book {
    let book = Book::from_file(asset(file)).unwrap();
    match locale {
        Some(tag) => book.with_locale(tag.parse().unwrap()),
        None => book,
    }
}

fn check(file: &str, book: &Book, options: &Options, spanish: bool) {
    for (formula, es, en) in CASES {
        let want = if spanish { es } else { en };
        assert_eq!(
            calculated(file, book, options, formula),
            json!(want),
            "{file}: {formula}"
        );
    }
}

#[test]
fn an_xlsx_calculates_with_the_locale_given_to_the_book() {
    let file = "semantics_es-CO.xlsx";
    check(file, &book(file, Some("es-CO")), &Options::default(), true);
    check(file, &book(file, Some("en-US")), &Options::default(), false);
}

#[test]
fn an_ods_calculates_with_the_locale_of_the_file() {
    // The ODS was saved by a `LibreOffice` with the `es-CO` regional setting, which writes it as the
    // language of the default cell style.
    let file = "semantics_es-CO.ods";
    check(file, &book(file, None), &Options::default(), true);
}

#[test]
fn the_locale_given_to_the_book_wins_over_the_one_of_the_file() {
    let file = "semantics_es-CO.ods";
    check(file, &book(file, Some("en-US")), &Options::default(), false);
}

#[test]
fn the_locale_of_a_call_wins_over_the_one_of_the_book() {
    let options = Options::from_json(&json!({ "locale": { "tag": "en-US" } })).unwrap();
    for file in ["semantics_es-CO.xlsx", "semantics_es-CO.ods"] {
        check(file, &book(file, Some("es-CO")), &options, false);
    }
    let options = Options::from_json(&json!({ "locale": { "tag": "es-CO" } })).unwrap();
    for file in ["semantics_es-CO.xlsx", "semantics_es-CO.ods"] {
        check(file, &book(file, Some("en-US")), &options, true);
    }
}

#[test]
fn the_locale_of_the_book_reaches_every_tool_of_the_compiler() {
    let file = "semantics_es-CO.xlsx";
    let book = book(file, Some("es-CO"));
    // `inspect` and `compile` report the locale the program is built for.
    let (program, _) = book.compile(&Options::default(), Mode::Values).unwrap();
    assert_eq!(program.build.options.locale, "es-CO");
    let (program, _) = Book::from_file(asset("semantics.xlsx"))
        .unwrap()
        .with_locale("de-DE".parse().unwrap())
        .compile(&Options::default(), Mode::Values)
        .unwrap();
    assert_eq!(program.build.options.locale, "de-DE");
    let inspection = book.inspect(&Options::default()).unwrap();
    assert!(inspection.diagnostics.iter().all(|d| !d.contains("locale")));
}

#[test]
fn a_locale_the_engine_does_not_have_is_an_error_that_names_it() {
    let file = "semantics.xlsx";
    let error = book(file, Some("fr-FR"))
        .compile(&Options::default(), Mode::Values)
        .map(|_| ())
        .unwrap_err()
        .to_string();
    assert!(error.contains("fr-FR"), "{error}");
}

/// `sheetc calculate` of `Numbers!C11` of `locale_es-CO` (`=TEXT(1234.5,"#,##0.00")`), made by a
/// `LibreOffice` with the `es-CO` regional setting, in an environment with no locale.
fn sheetc(file: &str, flags: &[&str]) -> String {
    use std::io::Write;
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-io/tests/assets")
        .join(file);
    let mut child = std::process::Command::new(env!("CARGO_BIN_EXE_sheetc"))
        .env_remove("LC_ALL")
        .env_remove("LC_CTYPE")
        .env_remove("LANG")
        .args(["calculate", "-i"])
        .arg(path)
        .args(["--request", "-"])
        .args(flags)
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .stderr(std::process::Stdio::piped())
        .spawn()
        .unwrap();
    child
        .stdin
        .take()
        .unwrap()
        .write_all(br#"{"ranges": ["Numbers!C11"]}"#)
        .unwrap();
    let out = child.wait_with_output().unwrap();
    assert!(
        out.status.success(),
        "{}",
        String::from_utf8_lossy(&out.stderr)
    );
    let answer: Json = serde_json::from_slice(&out.stdout).unwrap();
    answer["valueRanges"][0]["values"][0][0]
        .as_str()
        .unwrap()
        .to_string()
}

#[test]
fn sheetc_calculates_with_the_locale_of_the_workbook_and_the_flag_wins() {
    // An ODS says it; an XLSX has the system's, and this environment has none.
    assert_eq!(sheetc("locale_es-CO.ods", &[]), "1.234,50");
    assert_eq!(sheetc("locale_es-CO.xlsx", &[]), "1,234.50");
    assert_eq!(
        sheetc("locale_es-CO.xlsx", &["--locale", "es-CO"]),
        "1.234,50"
    );
    assert_eq!(
        sheetc("locale_es-CO.ods", &["--locale", "en-US"]),
        "1,234.50"
    );
}

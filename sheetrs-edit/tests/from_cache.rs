//! A workbook that depends on what the engine cannot calculate (another workbook, a macro, an extension) is run
//! as `LibreOffice` runs it, with no option (`sheetrs-compiler/docs/FIDELITY.md`, *What depends on what the engine
//! cannot calculate*, OBSERVED in 26.8 by `tests/assets/from_cache/probe.py`):
//!
//! * a cell that no edit reaches keeps the result the file caches, and the report holds a warning of its kind
//!   (`EXTERNAL_FROM_CACHE`, `MACRO_FROM_CACHE`, `EXTENSION_FROM_CACHE`);
//! * a cell that the run calculates (an edit reaches it, or every formula is calculated) is what `LibreOffice`
//!   gives: an add-in that is not installed (`_xll.NAME`) is `#NAME?`; a function only `LibreOffice` has
//!   (`WEBSERVICE`) has a result the engine cannot give, so it is an error that names the cause; a reference to
//!   another workbook keeps the stored value of the link, which is what the file caches;
//! * a workbook with macros is refused by the build, whatever the run, unless secure mode is on (`macro.*`): then
//!   what no edit reaches keeps the cache and `MACRO_FROM_CACHE` says so, and a calculated call is an error cell.
//!
//! The workbooks (`tests/assets/from_cache`, made by `make_fixtures.py` and `make_external.py`) hold, in `Calc`,
//! `A1` (the cell that depends on it, with a cached result), `B1 =A1+1` (cached 1000), `C1 =Data!D1*2` (cached 999,
//! it gives 10) and `E1 =C1+1` (cached 1000, it gives 11), in an XLSX and in an ODS. `A1` reads `Data!D1` in the
//! macro and the extension workbooks, so an edit of `D1` reaches it; `extension` has `F1 =WEBSERVICE("x")` too.

use serde_json::{Map, json};
use sheetrs_edit::{
    BookEdit, Calculation, EXTENSION_FROM_CACHE, EXTERNAL_FROM_CACHE, MACRO_FROM_CACHE,
    UNSUPPORTED_AS_ERROR,
};
use sheetrs_io::Book;
use sheetrs_io::reader::{CellValue, ExcelError};
use std::path::{Path, PathBuf};

/// What a kind of fixture is: its name, the code of its warning, what the error names when a calculated cell cannot be
/// answered, and what `A1` (and so `B1`) caches.
struct Kind {
    name: &'static str,
    code: &'static str,
    names: &'static [&'static str],
    a1: f64,
    /// The cell `C1` reads (`D1` of `Data`, of `Calc` in the workbook that reads another).
    input: &'static str,
}

const KINDS: [Kind; 3] = [
    Kind {
        name: "external",
        code: EXTERNAL_FROM_CACHE,
        names: &[],
        a1: 42.0,
        input: "Calc!D1",
    },
    Kind {
        name: "macro",
        code: MACRO_FROM_CACHE,
        names: &["macros", "secure mode"],
        a1: 77.0,
        input: "Data!D1",
    },
    Kind {
        name: "extension",
        code: EXTENSION_FROM_CACHE,
        names: &["does not implement", "WEBSERVICE"],
        a1: 77.0,
        input: "Data!D1",
    },
];

fn fixture(kind: &str, extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/from_cache")
        .join(format!("{kind}.{extension}"))
}

fn cell(bytes: &[u8], column: u32) -> CellValue {
    let workbook = Book::from_bytes(bytes.to_vec())
        .unwrap()
        .workbook()
        .unwrap();
    let sheet = workbook.sheets.iter().find(|s| s.name == "Calc").unwrap();
    sheet
        .cells
        .iter()
        .find(|(at, _)| (at.0, at.1) == (0, column))
        .map_or_else(
            || panic!("Calc row 1, column {column} has no value"),
            |(_, cell)| cell.value.clone(),
        )
}

fn number(bytes: &[u8], column: u32) -> f64 {
    match cell(bytes, column) {
        CellValue::Number(n) => n,
        other => panic!("Calc row 1, column {column} is {other:?}, not a number"),
    }
}

/// How a run is asked: what it does to the workbook.
#[derive(Clone, Copy)]
enum Run {
    /// Convert `Calc` to values; nothing reaches any formula.
    Convert,
    /// Convert `Calc` with every formula calculated (`calculate_all`).
    ConvertAll,
    /// Give the cell `C1` reads the value 6 in a `calculate` step: the formulas that read it are reached.
    Dirty,
    /// A `calculate` step with nothing given, in the mode of the cache.
    Calculate,
    /// A `calculate` step with nothing given and every formula calculated (`calculate_all`).
    CalculateAll,
}

fn run(
    kind: &Kind,
    extension: &str,
    how: Run,
) -> (Result<sheetrs_edit::Report, sheetrs_edit::Error>, Vec<u8>) {
    run_in_mode(kind, extension, how, false)
}

/// As [`run`], in secure mode or not.
fn run_in_mode(
    kind: &Kind,
    extension: &str,
    how: Run,
    secure: bool,
) -> (Result<sheetrs_edit::Report, sheetrs_edit::Error>, Vec<u8>) {
    let mut book = Book::from_file(fixture(kind.name, extension))
        .unwrap()
        .with_calculate_all(matches!(how, Run::ConvertAll | Run::CalculateAll))
        .with_secure_mode(secure);
    match how {
        Run::Convert | Run::ConvertAll => {
            book.formula_to_value(["Calc".parse().unwrap()]);
        }
        Run::Dirty => {
            let mut inputs = Map::new();
            inputs.insert(kind.input.into(), json!(6));
            book.calculate(Calculation {
                inputs,
                ..Calculation::default()
            });
        }
        Run::Calculate | Run::CalculateAll => {
            book.calculate(Calculation::default());
        }
    }
    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join(format!("out.{extension}"));
    let report = book.write(&out);
    let bytes = std::fs::read(&out).unwrap_or_default();
    (report, bytes)
}

fn codes(report: &sheetrs_edit::Report) -> Vec<&str> {
    report.diagnostics.iter().map(|d| d.code.as_str()).collect()
}

/// Nothing reaches the cells: they keep what the file caches (the XLSX, which `LibreOffice` shows as cached), in a
/// conversion and in a calculation that is given nothing. The report says it with the warning of the kind where the engine
/// runs (a calculation; an XLSX whose conversion has nothing stale converts the file's results and runs no engine).
#[test]
fn an_unreached_cell_keeps_its_cached_result_and_the_report_says_so() {
    for kind in &KINDS {
        for how in [Run::Convert, Run::Calculate] {
            // A workbook with macros is built only in secure mode.
            let (report, bytes) = run_in_mode(kind, "xlsx", how, kind.name == "macro");
            let report = report.unwrap_or_else(|e| panic!("{}: {e:#}", kind.name));
            if matches!(how, Run::Calculate) {
                assert_eq!(codes(&report), [kind.code], "{}", kind.name);
            } else {
                assert!(report.diagnostics.is_empty(), "{}", kind.name);
            }
            assert!(
                (number(&bytes, 0) - kind.a1).abs() < 1e-9,
                "{}: A1",
                kind.name
            );
            // `B1` reads `A1` and caches 1000: nothing reached it either.
            assert!(
                (number(&bytes, 1) - 1000.0).abs() < 1e-9,
                "{}: B1",
                kind.name
            );
            // The file caches 999 where the engine gives 10, and nothing reaches it.
            assert!(
                (number(&bytes, 2) - 999.0).abs() < 1e-9,
                "{}: C1 keeps the cache",
                kind.name
            );
        }
    }
}

/// A cell that depends on another workbook keeps the stored value of the link even when every formula is calculated
/// (`LibreOffice` `calculateAll()` gives the stored value, OBSERVED); what reads it is calculated from it.
#[test]
fn an_external_reference_keeps_the_stored_value_with_calculate_all() {
    let kind = &KINDS[0];
    for extension in ["xlsx", "ods"] {
        let (report, bytes) = run(kind, extension, Run::ConvertAll);
        let report = report.unwrap_or_else(|e| panic!("{extension}: {e:#}"));
        assert_eq!(codes(&report), [EXTERNAL_FROM_CACHE], "{extension}");
        assert!((number(&bytes, 0) - 42.0).abs() < 1e-9, "{extension}: A1");
        assert!((number(&bytes, 1) - 43.0).abs() < 1e-9, "{extension}: B1");
        // The rest is calculated (the file caches 999 and 1000 where it gives 10 and 11).
        assert!((number(&bytes, 2) - 10.0).abs() < 1e-9, "{extension}: C1");
        assert!((number(&bytes, 4) - 11.0).abs() < 1e-9, "{extension}: E1");
    }
}

/// A workbook with macros, and a function only `LibreOffice` has, in a cell that the run calculates: the result is not one the
/// engine can give, so the run is an error that names the cause (the cache is the result of a cell that is not calculated);
/// for the macros it is the refusal of the build, which names secure mode.
#[test]
fn a_macro_or_libreoffice_only_function_is_an_error_when_calculated() {
    for kind in &KINDS[1..] {
        for extension in ["xlsx", "ods"] {
            let (result, bytes) = run(kind, extension, Run::ConvertAll);
            let message = format!("{:#}", result.expect_err("an error, not a conversion"));
            assert!(
                bytes.is_empty(),
                "{}.{extension}: no file is left",
                kind.name
            );
            for name in kind.names {
                assert!(
                    message.contains(name),
                    "{}.{extension}: {name} in {message}",
                    kind.name
                );
            }
        }
    }
}

/// An edit that reaches the cell: an add-in that is not installed is `#NAME?` (OBSERVED, both formats), a workbook with macros is refused,
/// and the cells that no edit reaches (`WEBSERVICE`) keep what the file caches and are named by the warning.
#[test]
fn an_edit_that_reaches_a_cell_gives_what_libreoffice_gives_for_it() {
    // The add-in: `A1` reads `Data!D1`, so it is `#NAME?` and what reads it is too; `F1` is not reached.
    let kind = &KINDS[2];
    let (report, bytes) = run(kind, "xlsx", Run::Dirty);
    let report = report.unwrap_or_else(|e| panic!("{e:#}"));
    assert!(matches!(cell(&bytes, 0), CellValue::Error(_)), "A1");
    assert!(
        matches!(cell(&bytes, 1), CellValue::Error(_)),
        "B1 reads A1"
    );
    assert!((number(&bytes, 2) - 12.0).abs() < 1e-9, "C1 reads D1");
    assert!((number(&bytes, 4) - 13.0).abs() < 1e-9, "E1");
    assert_eq!(codes(&report), [EXTENSION_FROM_CACHE]);
    assert_eq!(report.diagnostics[0].context["cells"], "1", "F1 only");
    // The macro: the workbook is refused by the build, whatever the edit reaches.
    let (result, _) = run(&KINDS[1], "xlsx", Run::Dirty);
    let message = format!("{:#}", result.expect_err("an error"));
    assert!(message.contains("secure mode"), "{message}");
    // Another workbook: `A1` does not read `D1`, it keeps the stored value.
    let (report, bytes) = run(&KINDS[0], "xlsx", Run::Dirty);
    let report = report.unwrap_or_else(|e| panic!("{e:#}"));
    assert_eq!(codes(&report), [EXTERNAL_FROM_CACHE]);
    assert!((number(&bytes, 0) - 42.0).abs() < 1e-9);
    assert!((number(&bytes, 2) - 12.0).abs() < 1e-9);
}

#[test]
fn a_workbook_that_depends_on_nothing_is_calculated_with_no_warning() {
    let source = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/assets/cached/book.xlsx");
    let mut book = Book::from_file(source).unwrap();
    book.calculate(Calculation::default());
    let dir = tempfile::tempdir().unwrap();
    let report = book.write(dir.path().join("out.xlsx")).unwrap();
    assert!(report.diagnostics.is_empty());
}

/// What an error that a file cannot name is in a cell read back: an ODS keeps the code of `LibreOffice`, an XLSX holds
/// `#VALUE!` (as the export of `LibreOffice` does).
fn libre_error(extension: &str, code: u16) -> CellValue {
    CellValue::Error(if extension == "ods" {
        ExcelError::Libre(code)
    } else {
        ExcelError::Value
    })
}

/// In secure mode a macro that the run calculates is not an error of the run: the cell is the error `LibreOffice` shows
/// (`#MACRO?`, `Err:531`), what reads it is calculated from it, the rest of the workbook is calculated, and the report holds the warning.
#[test]
fn secure_mode_makes_a_calculated_macro_an_error_cell_and_warns() {
    for extension in ["xlsx", "ods"] {
        let (report, bytes) = run_in_mode(&KINDS[1], extension, Run::CalculateAll, true);
        let report = report.unwrap_or_else(|e| panic!("{extension}: {e:#}"));
        assert_eq!(
            cell(&bytes, 0),
            libre_error(extension, 531),
            "{extension}: A1"
        );
        assert_eq!(
            cell(&bytes, 1),
            libre_error(extension, 531),
            "{extension}: B1 reads A1"
        );
        assert!((number(&bytes, 2) - 10.0).abs() < 1e-9, "{extension}: C1");
        assert!((number(&bytes, 4) - 11.0).abs() < 1e-9, "{extension}: E1");
        assert_eq!(codes(&report), [UNSUPPORTED_AS_ERROR], "{extension}");
        let note = &report.diagnostics[0];
        assert_eq!(
            note.code, "UNSUPPORTED_AS_ERROR",
            "{extension}: the code is the contract"
        );
        assert_eq!(note.context["cells"], "1", "{extension}");
        assert!(
            note.message.contains("MYFUNC"),
            "{extension}: {}",
            note.message
        );
    }
}

/// The same for a function that only `LibreOffice` has: `WEBSERVICE` is `Err:540` (offline), and an add-in that is not
/// installed is `#NAME?`, which needs no warning because it is what `LibreOffice` gives.
#[test]
fn secure_mode_makes_a_calculated_libreoffice_only_function_an_error_cell_and_warns() {
    for extension in ["xlsx", "ods"] {
        let (report, bytes) = run_in_mode(&KINDS[2], extension, Run::CalculateAll, true);
        let report = report.unwrap_or_else(|e| panic!("{extension}: {e:#}"));
        assert_eq!(
            cell(&bytes, 5),
            libre_error(extension, 540),
            "{extension}: F1"
        );
        assert_eq!(
            cell(&bytes, 0),
            CellValue::Error(ExcelError::Name),
            "{extension}: A1 is the add-in"
        );
        assert!((number(&bytes, 2) - 10.0).abs() < 1e-9, "{extension}: C1");
        assert_eq!(codes(&report), [UNSUPPORTED_AS_ERROR], "{extension}");
        assert!(
            report.diagnostics[0].message.contains("WEBSERVICE"),
            "{extension}"
        );
    }
}

/// A cell that no edit reaches keeps what the file caches and the warning of its kind, in secure mode as by default.
#[test]
fn secure_mode_keeps_the_cache_of_a_cell_nothing_reaches() {
    let (report, bytes) = run_in_mode(&KINDS[2], "xlsx", Run::Dirty, true);
    let report = report.unwrap_or_else(|e| panic!("{e:#}"));
    assert_eq!(codes(&report), [EXTENSION_FROM_CACHE]);
    assert_eq!(
        cell(&bytes, 5),
        CellValue::Text("abc".into()),
        "F1 keeps its cache"
    );
}

/// An edit that reaches a macro in secure mode: the cell is the error and the rest is calculated from the new value.
#[test]
fn secure_mode_reaches_a_macro_through_an_edit() {
    for extension in ["xlsx", "ods"] {
        let (report, bytes) = run_in_mode(&KINDS[1], extension, Run::Dirty, true);
        let report = report.unwrap_or_else(|e| panic!("{extension}: {e:#}"));
        assert_eq!(
            cell(&bytes, 0),
            libre_error(extension, 531),
            "{extension}: A1"
        );
        assert!(
            (number(&bytes, 2) - 12.0).abs() < 1e-9,
            "{extension}: C1 reads D1"
        );
        assert_eq!(codes(&report), [UNSUPPORTED_AS_ERROR], "{extension}");
    }
}

/// The conversion of formulas to values is a run that calculates too: in secure mode it goes on with the macro and the
/// rest of `Calc` is converted from what the engine calculated.
#[test]
fn secure_mode_reaches_the_conversion_of_formulas_to_values() {
    for extension in ["xlsx", "ods"] {
        let (report, bytes) = run_in_mode(&KINDS[1], extension, Run::ConvertAll, true);
        let report = report.unwrap_or_else(|e| panic!("{extension}: {e:#}"));
        assert_eq!(codes(&report), [UNSUPPORTED_AS_ERROR], "{extension}");
        assert!((number(&bytes, 2) - 10.0).abs() < 1e-9, "{extension}: C1");
    }
}

/// The refusal is the build's, so every operation that builds the engine inherits it: a calculation with nothing given, every
/// formula calculated, an edit that reaches a formula, and a conversion of formulas to values that calculates. Each names
/// secure mode, leaves no file, and is the run of secure mode when it is on.
#[test]
fn every_run_that_builds_refuses_a_workbook_with_macros_unless_secure_mode_is_on() {
    for extension in ["xlsx", "ods"] {
        for how in [
            Run::Calculate,
            Run::CalculateAll,
            Run::Dirty,
            Run::ConvertAll,
        ] {
            let (result, bytes) = run(&KINDS[1], extension, how);
            let error = result.expect_err("refused");
            assert!(
                format!("{error:#}").contains("the workbook was refused"),
                "{extension}: {error:#}"
            );
            assert!(format!("{error:#}").contains("secure mode"), "{extension}");
            assert!(bytes.is_empty(), "{extension}: no file is left");
            let (secure, bytes) = run_in_mode(&KINDS[1], extension, how, true);
            secure.unwrap_or_else(|e| panic!("{extension}: {e:#}"));
            assert!(!bytes.is_empty(), "{extension}");
        }
    }
}

/// The value ranges of a calculation build the program too.
#[test]
fn the_value_ranges_of_a_workbook_with_macros_are_refused_unless_secure_mode_is_on() {
    for extension in ["xlsx", "ods"] {
        let ranges = |secure: bool| {
            let mut book = Book::from_file(fixture("macro", extension))
                .unwrap()
                .with_secure_mode(secure);
            book.calculate(Calculation {
                ranges: Some(vec!["Calc!A1:E1".into()]),
                answer: true,
                ..Calculation::default()
            });
            book.value_ranges()
        };
        let error = ranges(false).expect_err("refused");
        assert!(format!("{error:#}").contains("secure mode"), "{extension}");
        ranges(true).unwrap_or_else(|e| panic!("{extension}: {e:#}"));
    }
}

/// A conversion of an XLSX whose results are all right runs no engine, so nothing builds and nothing refuses: the file's
/// results are converted as they are (the macros are not run either way).
#[test]
fn an_xlsx_conversion_that_needs_no_engine_does_not_build_and_is_not_refused() {
    let (report, bytes) = run(&KINDS[1], "xlsx", Run::Convert);
    report.unwrap();
    assert!((number(&bytes, 0) - 77.0).abs() < 1e-9);
}

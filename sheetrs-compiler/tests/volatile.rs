//! `now` and `seed` fix `NOW()`, `TODAY()` and `RAND()` in every tool of `sheetc` that builds or
//! calculates, for an XLSX and for an ODS: the flags of each command (`build`, `calculate`,
//! `evaluate`, `verify`, and an executable's own), the keys of `sheetc.toml`, and, in memory, the
//! `Book` (`Book::with_volatile`) with the options of a call winning over it.
//!
//! The workbooks (`tests/assets/volatile.{xlsx,ods}`, made by `sheetrs-edit/scripts/gen_volatile_fixtures.py`)
//! hold `V!A1 =NOW()`, `V!B1 =TODAY()` and `V!C1 =RAND()`.

use serde_json::{Value as Json, json};
use sheetrs_compiler::memory::{BookCompile, Options};
use sheetrs_engine::ir::Mode;
use sheetrs_io::{Book, Volatile};
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};

const SHEETC: &str = env!("CARGO_BIN_EXE_sheetc");
const EXTENSIONS: [&str; 2] = ["xlsx", "ods"];
/// 2024-03-15 18:00.
const NOW: f64 = 45366.75;

fn asset(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(format!("volatile.{extension}"))
}

fn run(program: &Path, args: &[&str]) -> (i32, String, String) {
    let out = Command::new(program)
        .args(args)
        .stdin(Stdio::null())
        .output()
        .unwrap();
    (
        out.status.code().unwrap_or(-1),
        String::from_utf8_lossy(&out.stdout).into_owned(),
        String::from_utf8_lossy(&out.stderr).into_owned(),
    )
}

/// `[NOW(), TODAY(), RAND()]` of the answer of a command or an executable (the range `V!A1:C1`).
fn triple(answer: &str) -> [f64; 3] {
    let json: Json = serde_json::from_str(answer).unwrap_or_else(|e| panic!("{answer}: {e}"));
    let row = &json["valueRanges"][0]["values"][0];
    [0, 1, 2].map(|i| row[i].as_f64().unwrap_or_else(|| panic!("{answer}")))
}

fn calculate(extension: &str, flags: &[&str]) -> [f64; 3] {
    let file = asset(extension);
    let mut args = vec![
        "calculate",
        "-i",
        file.to_str().unwrap(),
        "--range",
        "V!A1:C1",
    ];
    args.extend_from_slice(flags);
    let (code, out, err) = run(Path::new(SHEETC), &args);
    assert_eq!(code, 0, "{err}");
    triple(&out)
}

fn same(a: f64, b: f64) -> bool {
    a.to_bits() == b.to_bits()
}

#[test]
fn calculate_takes_now_and_seed() {
    for extension in EXTENSIONS {
        let [now, today, rand] = calculate(extension, &["--now", "45366.75", "--seed", "42"]);
        assert!((now - NOW).abs() < 1e-9, "{extension}: {now}");
        assert!((today - NOW.floor()).abs() < 1e-9, "{extension}: {today}");
        let [_, _, again] = calculate(extension, &["--seed", "42"]);
        let [_, _, other] = calculate(extension, &["--seed", "44"]);
        assert!(same(rand, again), "{extension}");
        assert!(!same(rand, other), "{extension}");
    }
}

#[test]
fn a_configuration_file_fixes_now_and_seed_and_a_flag_wins_over_it() {
    for extension in EXTENSIONS {
        let dir = tempfile::tempdir().unwrap();
        let file = dir.path().join(format!("book.{extension}"));
        std::fs::copy(asset(extension), &file).unwrap();
        std::fs::write(
            dir.path().join("sheetc.toml"),
            "now = 45366.75\nseed = 42\n",
        )
        .unwrap();
        let args = |flags: &[&str]| {
            let mut args = vec![
                "calculate",
                "-i",
                file.to_str().unwrap(),
                "--range",
                "V!A1:C1",
            ];
            args.extend_from_slice(flags);
            let (code, out, err) = run(Path::new(SHEETC), &args);
            assert_eq!(code, 0, "{err}");
            triple(&out)
        };
        let [now, _, rand] = args(&[]);
        assert!((now - NOW).abs() < 1e-9, "{extension}: {now}");
        let [_, _, by_flag] = args(&["--seed", "42"]);
        assert!(
            same(rand, by_flag),
            "{extension}: the file's seed is --seed 42"
        );
        let [now, _, other] = args(&["--now", "40000", "--seed", "44"]);
        assert!((now - 40000.0).abs() < 1e-9, "{extension}: {now}");
        assert!(!same(rand, other), "{extension}");
    }
}

#[test]
fn evaluate_takes_now_and_seed() {
    let (code, out, err) = run(
        Path::new(SHEETC),
        &["evaluate", "=NOW()", "--now", "45366.75"],
    );
    assert_eq!(code, 0, "{err}");
    assert_eq!(out.trim(), "45366.75");
    let rand = |seed: &str| {
        let (code, out, err) = run(Path::new(SHEETC), &["evaluate", "=RAND()", "--seed", seed]);
        assert_eq!(code, 0, "{err}");
        out
    };
    assert_eq!(rand("42"), rand("42"));
    assert_ne!(rand("42"), rand("44"));
}

#[test]
fn verify_takes_now_and_seed() {
    for extension in EXTENSIONS {
        let file = asset(extension);
        let (_, out, err) = run(
            Path::new(SHEETC),
            &[
                "verify",
                "-i",
                file.to_str().unwrap(),
                "--now",
                "45366.75",
                "--seed",
                "7",
            ],
        );
        assert!(!err.contains("unknown option"), "{extension}: {err}");
        assert!(out.contains("volatile"), "{extension}: {out}");
    }
}

/// The executable of the workbook, built with `build_flags`.
fn executable(extension: &str, dir: &Path, build_flags: &[&str]) -> PathBuf {
    let exe = dir.join(format!("volatile-{extension}"));
    let file = asset(extension);
    let mut args = vec![
        "build",
        "-i",
        file.to_str().unwrap(),
        "--runtime",
        SHEETC,
        "-o",
        exe.to_str().unwrap(),
    ];
    args.extend_from_slice(build_flags);
    let (code, _, err) = run(Path::new(SHEETC), &args);
    assert_eq!(code, 0, "{err}");
    exe
}

fn answer(exe: &Path, flags: &[&str]) -> [f64; 3] {
    let mut args = vec!["--range", "V!A1:C1"];
    args.extend_from_slice(flags);
    let (code, out, err) = run(exe, &args);
    assert_eq!(code, 0, "{err}");
    triple(&out)
}

#[test]
fn an_executable_takes_now_and_seed_at_run_time() {
    let dir = tempfile::tempdir().unwrap();
    for extension in EXTENSIONS {
        let exe = executable(extension, dir.path(), &[]);
        let [now, today, rand] = answer(&exe, &["--now", "45366.75", "--seed", "42"]);
        assert!((now - NOW).abs() < 1e-9, "{extension}: {now}");
        assert!((today - NOW.floor()).abs() < 1e-9, "{extension}: {today}");
        assert!(
            same(rand, answer(&exe, &["--seed", "42"])[2]),
            "{extension}"
        );
        assert!(
            !same(rand, answer(&exe, &["--seed", "44"])[2]),
            "{extension}"
        );
    }
}

#[test]
fn the_build_fixes_now_and_seed_and_the_run_flags_win_over_them() {
    let dir = tempfile::tempdir().unwrap();
    for extension in EXTENSIONS {
        let exe = executable(
            extension,
            dir.path(),
            &["--now", "45366.75", "--seed", "42"],
        );
        let [now, _, rand] = answer(&exe, &[]);
        assert!((now - NOW).abs() < 1e-9, "{extension}: {now}");
        let [now, _, other] = answer(&exe, &["--now", "40000", "--seed", "44"]);
        assert!((now - 40000.0).abs() < 1e-9, "{extension}: {now}");
        assert!(!same(rand, other), "{extension}");
        // The same seed at the build and at the run is the same stream.
        let plain = executable(extension, dir.path(), &[]);
        assert!(
            same(rand, answer(&plain, &["--seed", "42"])[2]),
            "{extension}"
        );
    }
}

#[test]
fn an_embedded_executable_writes_a_workbook_calculated_with_its_now_and_seed() {
    let dir = tempfile::tempdir().unwrap();
    for extension in EXTENSIONS {
        let exe = executable(extension, dir.path(), &["--embed-template"]);
        let out = dir.path().join(format!("out.{extension}"));
        let (code, json, err) = run(
            &exe,
            &[
                "-o",
                out.to_str().unwrap(),
                "--json",
                "--range",
                "V!A1:C1",
                "--now",
                "45366.75",
                "--seed",
                "42",
            ],
        );
        assert_eq!(code, 0, "{err}");
        let [now, _, rand] = triple(&json);
        assert!((now - NOW).abs() < 1e-9, "{extension}: {now}");
        assert!(
            same(rand, calculate(extension, &["--seed", "42"])[2]),
            "{extension}"
        );
        // What the workbook caches is what the executable calculated.
        let written = Book::from_file(&out).unwrap().workbook().unwrap();
        let sheet = written.sheets.iter().find(|s| s.name == "V").unwrap();
        let cached = |col: u32| match sheet
            .cells
            .iter()
            .find(|(at, _)| (at.0, at.1) == (0, col))
            .map(|(_, cell)| cell.value.clone())
        {
            Some(sheetrs_io::reader::CellValue::Number(n)) => n,
            other => panic!("{extension}: {other:?}"),
        };
        assert!((cached(0) - NOW).abs() < 1e-9, "{extension}");
        assert!(same(cached(2), rand), "{extension}");
    }
}

fn book(extension: &str) -> Book {
    Book::from_file(asset(extension)).unwrap()
}

fn values(book: &Book, options: &Options) -> [f64; 3] {
    let done = book
        .calculate(
            &json!({"ranges": ["V!A1:C1"]}),
            options,
            Mode::Values,
            false,
        )
        .unwrap();
    let json = done.answer.unwrap().to_json();
    let row = &json["valueRanges"][0]["values"][0];
    [0, 1, 2].map(|i| row[i].as_f64().unwrap())
}

#[test]
fn the_volatile_of_a_book_fixes_the_calculation_in_memory() {
    for extension in EXTENSIONS {
        let given = book(extension).with_volatile(Volatile {
            now: Some(NOW),
            seed: Some(42),
        });
        let [now, today, rand] = values(&given, &Options::default());
        assert!((now - NOW).abs() < 1e-9, "{extension}: {now}");
        assert!((today - NOW.floor()).abs() < 1e-9, "{extension}: {today}");
        assert!(
            same(rand, calculate(extension, &["--seed", "42"])[2]),
            "{extension}"
        );
        // Without them the clock decides: a time later than 2024-01-01 that is not the fixed one.
        let [clock, _, _] = values(&book(extension), &Options::default());
        assert!(
            clock > 45292.0 && (clock - NOW).abs() > 1e-6,
            "{extension}: {clock}"
        );
    }
}

#[test]
fn the_options_of_a_call_win_over_the_volatile_of_the_book_field_by_field() {
    for extension in EXTENSIONS {
        let given = book(extension).with_volatile(Volatile {
            now: Some(40000.0),
            seed: Some(42),
        });
        let mut options = Options::default();
        options.config.now = Some(NOW);
        let [now, _, rand] = values(&given, &options);
        assert!((now - NOW).abs() < 1e-9, "{extension}: {now}");
        // The seed was left open by the call: it is the book's.
        assert!(
            same(rand, calculate(extension, &["--seed", "42"])[2]),
            "{extension}"
        );
        options.config.seed = Some(44);
        let [_, _, rand] = values(&given, &options);
        assert!(
            same(rand, calculate(extension, &["--seed", "44"])[2]),
            "{extension}"
        );
    }
}

#[test]
fn the_build_of_a_book_takes_its_volatile_too() {
    for extension in EXTENSIONS {
        let given = book(extension).with_volatile(Volatile {
            now: Some(NOW),
            seed: Some(42),
        });
        let (program, _) = given.compile(&Options::default(), Mode::Values).unwrap();
        assert_eq!(program.options.now, Some(NOW), "{extension}");
        assert_eq!(program.options.rand_seed, Some(42), "{extension}");
    }
}

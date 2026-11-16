//! `now` and `seed` of a `calculate` step fix `NOW()`, `TODAY()` and `RAND()` in an XLSX and in an ODS,
//! and the recalculation the planner adds after an edit runs with them too. The workbooks
//! (`tests/assets/volatile`, made by `scripts/gen_volatile_fixtures.py`) hold `V!A1 =NOW()`,
//! `V!B1 =TODAY()` and `V!C1 =RAND()`, and rows to edit.

use sheetrs_edit::{BookEdit, Calculation};
use sheetrs_io::reader::CellValue;
use sheetrs_io::{Book, Volatile};
use std::path::{Path, PathBuf};

const EXTENSIONS: [&str; 2] = ["xlsx", "ods"];
/// 2024-03-15 18:00 (serial 45366.75).
const NOW: f64 = 45366.75;

fn fixture(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/volatile")
        .join(format!("volatile.{extension}"))
}

fn number(bytes: &[u8], row: u32, column: u32) -> f64 {
    let workbook = Book::from_bytes(bytes.to_vec())
        .unwrap()
        .workbook()
        .unwrap();
    let sheet = workbook.sheets.iter().find(|s| s.name == "V").unwrap();
    match sheet
        .cells
        .iter()
        .find(|(at, _)| (at.0, at.1) == (row, column))
        .map(|(_, cell)| cell.value.clone())
    {
        Some(CellValue::Number(n)) => n,
        other => panic!("V {row},{column} is {other:?}, not a number"),
    }
}

/// `NOW()`, `TODAY()` and `RAND()` of the workbook after the plan of `book`.
fn volatile(book: &mut Book) -> (f64, f64, f64) {
    let dir = tempfile::tempdir().unwrap();
    let out = dir
        .path()
        .join(format!("out.{}", book.format().extension()));
    book.write(&out).unwrap();
    let bytes = std::fs::read(out).unwrap();
    (
        number(&bytes, 0, 0),
        number(&bytes, 0, 1),
        number(&bytes, 0, 2),
    )
}

fn calculate(extension: &str, now: Option<f64>, seed: Option<u64>) -> (f64, f64, f64) {
    let mut book = Book::from_file(fixture(extension)).unwrap();
    book.calculate(Calculation {
        now,
        seed,
        ..Calculation::default()
    });
    volatile(&mut book)
}

#[test]
fn now_fixes_now_and_today() {
    for extension in EXTENSIONS {
        let (now, today, _) = calculate(extension, Some(NOW), None);
        assert!((now - NOW).abs() < 1e-9, "{extension}: {now}");
        assert!((today - NOW.floor()).abs() < 1e-9, "{extension}: {today}");
    }
}

#[test]
fn seed_fixes_rand() {
    for extension in EXTENSIONS {
        let (_, _, first) = calculate(extension, None, Some(42));
        let (_, _, again) = calculate(extension, None, Some(42));
        let (_, _, other) = calculate(extension, None, Some(44));
        assert!((0.0..1.0).contains(&first), "{extension}: {first}");
        assert!(first.to_bits() == again.to_bits(), "{extension}");
        assert!(first.to_bits() != other.to_bits(), "{extension}");
    }
}

#[test]
fn an_even_seed_and_the_next_odd_one_give_distinct_streams() {
    for extension in EXTENSIONS {
        for even in [0, 42, 1000] {
            let (_, _, a) = calculate(extension, None, Some(even));
            let (_, _, b) = calculate(extension, None, Some(even + 1));
            assert!(
                a.to_bits() != b.to_bits(),
                "{extension}: seeds {even} and {}",
                even + 1
            );
        }
    }
}

#[test]
fn without_now_and_seed_the_clock_is_the_systems() {
    for extension in EXTENSIONS {
        let (now, _, _) = calculate(extension, None, None);
        // Later than 2024-01-01: it is not a value the workbook or a test fixed.
        assert!(now > 45292.0, "{extension}: {now}");
        assert!((now - NOW).abs() > 1e-6, "{extension}: {now}");
    }
}

#[test]
fn the_recalculation_after_a_deletion_runs_with_the_clock_of_the_calculate_step() {
    for extension in EXTENSIONS {
        let mut book = Book::from_file(fixture(extension)).unwrap();
        book.calculate(Calculation {
            now: Some(NOW),
            seed: Some(7),
            ..Calculation::default()
        })
        .delete_rows(["V!3".parse().unwrap()]);
        let (now, today, rand) = volatile(&mut book);
        assert!((now - NOW).abs() < 1e-9, "{extension}: {now}");
        assert!((today - NOW.floor()).abs() < 1e-9, "{extension}: {today}");
        // The same seed calculated without a deletion gives the same number.
        let (_, _, alone) = calculate(extension, Some(NOW), Some(7));
        assert!(rand.to_bits() == alone.to_bits(), "{extension}");
    }
}

#[test]
fn the_volatile_of_the_book_fixes_the_recalculation_an_edit_leaves() {
    for extension in EXTENSIONS {
        let mut book = Book::from_file(fixture(extension))
            .unwrap()
            .with_volatile(Volatile {
                now: Some(NOW),
                seed: Some(7),
            });
        book.delete_rows(["V!3".parse().unwrap()]);
        let (now, today, rand) = volatile(&mut book);
        assert!((now - NOW).abs() < 1e-9, "{extension}: {now}");
        assert!((today - NOW.floor()).abs() < 1e-9, "{extension}: {today}");
        let (_, _, alone) = calculate(extension, Some(NOW), Some(7));
        assert!(rand.to_bits() == alone.to_bits(), "{extension}");
    }
}

#[test]
fn the_now_of_a_step_wins_over_the_one_of_the_book() {
    for extension in EXTENSIONS {
        let mut book = Book::from_file(fixture(extension))
            .unwrap()
            .with_volatile(Volatile {
                now: Some(40000.0),
                seed: None,
            });
        book.calculate(Calculation {
            now: Some(NOW),
            ..Calculation::default()
        });
        let (now, _, _) = volatile(&mut book);
        assert!((now - NOW).abs() < 1e-9, "{extension}: {now}");
    }
}

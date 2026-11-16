//! What the writers change of a result when they write it, against `LibreOffice` 26.8 (`sheetrs-io/README.md`, *What the
//! writers change of a result when they write it*).
//!
//! The reference is the file `LibreOffice` saves and opens again (`CONTRIBUTING.md`). `tests/assets/writer_results/make_oracle.py`
//! makes `kinds.{xlsx,ods}` (and `kinds_1904.*`, in the 1904 date system) with one formula per kind of result in the sheet `K`
//! (numbers, texts, booleans, every error, dates, times, percentages, currencies, an array) and one constant per kind in the sheet
//! `C`, converts the formulas to values with `LibreOffice` (*Data > Calculate > Formula to Value*), saves, opens again and records
//! what it shows (`lo`); then it runs `sheetrs edit` on the same file and records what `LibreOffice` shows of that one (`default`,
//! `all`, `recalc`, `sheetrs_constants`), with `differences_*`: the cells where the two are not the same, in the value, the type,
//! the formula or the number format.
//!
//! This test checks that
//!
//! * the library gives, for every kind, the value `LibreOffice` shows (the conversion in the default mode and with `calculate_all`,
//!   the results written next to their formulas, and the constants), the numbers within 15 significant digits;
//! * the cells where the writers differ from `LibreOffice` are exactly the recorded ones, which are in
//!   `docs/approved-differences.md` as proposed: a change that fixes one, or adds one, is seen here.

use serde_json::Value as Json;
use sheetrs_edit::{BookEdit, CellWrites, Constant};
use sheetrs_io::Book;
use sheetrs_io::reader::{CellValue, Workbook};
use std::path::{Path, PathBuf};
use std::sync::LazyLock;

static ORACLE: LazyLock<Json> = LazyLock::new(|| {
    let path = dir().join("oracle.json");
    serde_json::from_str(&std::fs::read_to_string(path).unwrap()).unwrap()
});

fn dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/assets/writer_results")
}

const VARIANTS: [&str; 4] = ["xlsx", "ods", "xlsx_1904", "ods_1904"];

fn extension(variant: &str) -> &str {
    variant.split('_').next().unwrap()
}

fn fixture(variant: &str) -> PathBuf {
    let (ext, rest) = variant
        .split_once('_')
        .map_or((variant, ""), |(e, r)| (e, r));
    let suffix = if rest.is_empty() {
        String::new()
    } else {
        format!("_{rest}")
    };
    dir().join(format!("kinds{suffix}.{ext}"))
}

/// What a cell holds, as far as this test tells.
#[derive(Debug, Clone, PartialEq)]
enum Seen {
    Empty,
    Number(f64),
    Text(String),
    Boolean(bool),
    Error,
}

impl Seen {
    /// The same, with the numbers equal to 15 significant digits (what `LibreOffice` keeps of a number it writes).
    fn same(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Number(a), Self::Number(b)) => {
                a == b
                    || (a.is_nan() && b.is_nan())
                    || (a - b).abs() <= 1e-14 * a.abs().max(b.abs())
            }
            (a, b) => a == b,
        }
    }
}

fn number_of(value: &Json) -> f64 {
    match value {
        Json::String(text) => match text.as_str() {
            "nan" => f64::NAN,
            "inf" => f64::INFINITY,
            "-inf" => f64::NEG_INFINITY,
            other => other.parse().unwrap(),
        },
        other => other.as_f64().unwrap(),
    }
}

/// What `LibreOffice` shows of a cell, from the oracle.
fn libreoffice(cell: &Json) -> Seen {
    let logical = cell["format"]["type"].as_i64().unwrap() & 1024 != 0;
    let number = || {
        let n = number_of(&cell["value"]);
        if logical {
            Seen::Boolean(n != 0.0)
        } else {
            Seen::Number(n)
        }
    };
    match cell["kind"].as_str().unwrap() {
        "EMPTY" => Seen::Empty,
        "TEXT" => Seen::Text(cell["string"].as_str().unwrap().to_string()),
        "VALUE" => number(),
        _ if cell["error"].as_i64().unwrap() != 0 => Seen::Error,
        _ => match cell["formula"].as_str().unwrap() {
            "=TRUE()" => Seen::Boolean(true),
            "=FALSE()" => Seen::Boolean(false),
            _ if cell["result"].as_i64().unwrap() == 2 => {
                Seen::Text(cell["string"].as_str().unwrap().to_string())
            }
            _ => number(),
        },
    }
}

fn seen(book: &Workbook, sheet: &str, row: u32) -> Seen {
    let Some(cell) = book
        .sheets
        .iter()
        .find(|s| s.name == sheet)
        .unwrap()
        .cells
        .get(&(row, 0))
    else {
        return Seen::Empty;
    };
    match &cell.value {
        CellValue::Empty => Seen::Empty,
        CellValue::Number(n) => Seen::Number(*n),
        CellValue::Text(t) => Seen::Text(t.to_string()),
        CellValue::Boolean(b) => Seen::Boolean(*b),
        CellValue::Error(_) => Seen::Error,
    }
}

fn labels(key: &str) -> Vec<String> {
    ORACLE[key]
        .as_array()
        .unwrap()
        .iter()
        .map(|l| l.as_str().unwrap().to_string())
        .collect()
}

/// The workbook the library writes for `plan`, read again.
fn written(variant: &str, plan: impl FnOnce(&mut Book)) -> Workbook {
    let mut book = Book::from_file(fixture(variant)).unwrap();
    plan(&mut book);
    let dir = tempfile::tempdir().unwrap();
    let out = dir.path().join(format!("out.{}", extension(variant)));
    book.write(&out).unwrap();
    Book::from_file(&out).unwrap().workbook().unwrap()
}

/// The cells of `K` where the library's file is not what `LibreOffice` shows (`expected`: the key of the oracle).
fn mismatches(variant: &str, book: &Workbook, expected: &str) -> Vec<String> {
    let names = labels("formulas");
    let cells = ORACLE[variant][expected].as_array().unwrap();
    (0..cells.len())
        .filter(|&row| {
            !seen(book, "K", u32::try_from(row).unwrap()).same(&libreoffice(&cells[row]))
        })
        .map(|row| {
            names
                .get(row)
                .cloned()
                .unwrap_or_else(|| "array (second cell)".into())
        })
        .collect()
}

#[test]
fn a_conversion_gives_the_values_libreoffice_gives_in_the_default_mode() {
    for variant in VARIANTS {
        let book = written(variant, |book| {
            book.formula_to_value(["K".parse().unwrap()]);
        });
        let wrong = mismatches(variant, &book, "lo");
        assert!(wrong.is_empty(), "{variant}: {wrong:?}");
    }
}

#[test]
fn a_conversion_gives_the_values_libreoffice_gives_with_calculate_all() {
    for variant in VARIANTS {
        let book = written(variant, |book| {
            *book = Book::from_file(fixture(variant))
                .unwrap()
                .with_calculate_all(true);
            book.formula_to_value(["K".parse().unwrap()]);
        });
        let wrong = mismatches(variant, &book, "lo");
        assert!(wrong.is_empty(), "{variant}: {wrong:?}");
    }
}

#[test]
fn the_results_written_next_to_the_formulas_are_the_ones_libreoffice_gives() {
    for variant in VARIANTS {
        let book = written(variant, |book| {
            *book = Book::from_file(fixture(variant))
                .unwrap()
                .with_calculate_all(true);
            book.calculate(sheetrs_edit::Calculation::default());
        });
        let wrong = mismatches(variant, &book, "recalc_lo");
        assert!(wrong.is_empty(), "{variant}: {wrong:?}");
    }
}

/// An error is not kept by a conversion: the cell is the empty text, in both formats, as `LibreOffice` leaves it (`Data > Calculate >
/// Formula to Value` of a formula that gives an error), for every kind of error, in the two modes.
#[test]
fn a_conversion_leaves_an_empty_text_where_the_formula_gave_an_error() {
    let names = labels("formulas");
    let errors: Vec<usize> = names
        .iter()
        .enumerate()
        .filter(|(_, name)| {
            name.starts_with('#') && !name.starts_with("#DIV/0! (")
                || name.starts_with("Err:")
                || name.contains("error")
        })
        .filter(|(_, name)| !name.contains("text that looks"))
        .map(|(i, _)| i)
        .collect();
    assert_eq!(errors.len(), 11, "{errors:?}");
    for variant in VARIANTS {
        for calculate_all in [false, true] {
            let book = written(variant, |book| {
                *book = Book::from_file(fixture(variant))
                    .unwrap()
                    .with_calculate_all(calculate_all);
                book.formula_to_value(["K".parse().unwrap()]);
            });
            for &row in &errors {
                let got = seen(&book, "K", u32::try_from(row).unwrap());
                // `LibreOffice` shows the same, and not an error: an empty cell in an XLSX, an empty text in an ODS.
                assert_eq!(
                    got,
                    libreoffice(&ORACLE[variant]["lo"][row]),
                    "{variant} {calculate_all}: {}",
                    names[row]
                );
                let empty = if extension(variant) == "ods" {
                    Seen::Text(String::new())
                } else {
                    Seen::Empty
                };
                assert_eq!(got, empty, "{variant} {calculate_all}: {}", names[row]);
            }
        }
    }
}

/// The cells of `C` (`constants_spec` of the oracle) written as constants, over cells that exist in `W`.
fn constants(variant: &str) -> Workbook {
    written(variant, |book| {
        let mut writes = CellWrites::new();
        for (row, spec) in ORACLE["constants_spec"]
            .as_array()
            .unwrap()
            .iter()
            .enumerate()
        {
            let row = u32::try_from(row).unwrap();
            let constant = match spec["kind"].as_str().unwrap() {
                "number" => Constant::Number(number_of(&spec["value"])),
                "text" => Constant::Text(spec["value"].as_str().unwrap().to_string()),
                "boolean" => Constant::Boolean(spec["value"].as_bool().unwrap()),
                _ => Constant::Empty,
            };
            writes.constant("W", row, 0, constant);
        }
        book.write_cells(writes);
    })
}

#[test]
fn the_constants_written_are_the_ones_libreoffice_saves() {
    let names = labels("constants");
    for variant in VARIANTS {
        let book = constants(variant);
        for (row, name) in names.iter().enumerate() {
            let want = libreoffice(&ORACLE[variant]["source_constants"][row]);
            let got = seen(&book, "W", u32::try_from(row).unwrap());
            // An empty text constant is an empty cell for `LibreOffice` and, in an ODS, a text cell with no characters here: proposed.
            let proposed = extension(variant) == "ods" && name == "empty text";
            // `LibreOffice` writes 0 for NaN and the infinities in an ODS, and shows 0 for what is written here too (`NaN`, `inf`: the
            // oracle records the same value for both, `differences_constants`); the reader of this crate reads the text as it is.
            let not_finite =
                extension(variant) == "ods" && ["NaN", "+inf", "-inf"].contains(&name.as_str());
            assert!(
                got.same(&want) || proposed || not_finite,
                "{variant}: {name}: {got:?} is not {want:?}"
            );
        }
    }
}

fn differing(variant: &str, key: &str, names: &[String]) -> Vec<String> {
    ORACLE[variant][key]
        .as_array()
        .unwrap()
        .iter()
        .map(|i| names[usize::try_from(i.as_u64().unwrap()).unwrap()].clone())
        .collect()
}

fn strings(of: &[&str]) -> Vec<String> {
    of.iter().map(|s| (*s).to_string()).collect()
}

/// `LibreOffice` writes the numbers with 15 significant digits; the writers here write all the digits of the value (the engine's
/// result, or the constant given). The values differ in the digits after the 15th, and `LibreOffice` shows the same text.
/// Proposed in `docs/approved-differences.md`.
#[test]
fn numbers_keep_all_their_digits_where_libreoffice_keeps_fifteen() {
    let names = labels("formulas");
    for variant in VARIANTS {
        let (xlsx, all) = (
            extension(variant) == "xlsx",
            strings(&["fraction", "float sum 0.1+0.2", "2^53+1"]),
        );
        let mut with_engine = all.clone();
        if xlsx {
            with_engine.push("date and time".into());
        }
        // An XLSX converts what the file caches (15 digits already) in the default mode; an ODS is always calculated.
        let default = if xlsx { Vec::new() } else { all };
        assert_eq!(
            differing(variant, "differences_default", &names),
            default,
            "{variant}"
        );
        assert_eq!(
            differing(variant, "differences_all", &names),
            with_engine,
            "{variant}"
        );
        if xlsx {
            assert_eq!(
                differing(variant, "differences_recalc", &names),
                with_engine,
                "{variant}"
            );
        } else {
            assert!(
                differing(variant, "differences_recalc", &names).is_empty(),
                "{variant}"
            );
        }
    }
    // The value the writer keeps is the whole double, and the one `LibreOffice` keeps is the one it rounded.
    let book = written("xlsx", |book| {
        *book = Book::from_file(fixture("xlsx"))
            .unwrap()
            .with_calculate_all(true);
        book.formula_to_value(["K".parse().unwrap()]);
    });
    assert_eq!(seen(&book, "K", 1), Seen::Number(1.0 / 3.0));
    assert_eq!(
        libreoffice(&ORACLE["xlsx"]["lo"][1]),
        Seen::Number(0.333_333_333_333_333)
    );
}

/// What differs in the constants (proposed): the numbers of more than 15 digits, the text that `LibreOffice` marks with the text format
/// when it would be read as a number or a formula (the value is the same), the empty text of an ODS, and the infinities of an XLSX
/// (`inf` here, `INF` in `LibreOffice`, which reads `inf` as 0).
#[test]
fn the_constants_that_differ_from_libreoffice_are_the_recorded_ones() {
    let names = labels("constants");
    let common = [
        "123456789012345678",
        "2^53+1 as a double",
        "leading =",
        "numeric text",
        "text TRUE",
    ];
    for variant in VARIANTS {
        let mut want = strings(&common);
        if extension(variant) == "xlsx" {
            want.extend(strings(&["+inf", "-inf"]));
        } else {
            want.insert(2, "empty text".into());
        }
        assert_eq!(
            differing(variant, "differences_constants", &names),
            want,
            "{variant}"
        );
    }
}

/// Nothing else differs: a boolean, a date, a time, a percentage, a currency, a text with spaces, newlines or unicode, an error, the
/// largest and the smallest numbers are the same for `LibreOffice` after the conversion.
#[test]
fn every_other_kind_is_the_same_for_libreoffice() {
    let names = labels("formulas");
    for variant in VARIANTS {
        for key in [
            "differences_default",
            "differences_all",
            "differences_recalc",
        ] {
            for name in differing(variant, key, &names) {
                assert!(
                    ["fraction", "float sum 0.1+0.2", "2^53+1", "date and time"]
                        .contains(&name.as_str()),
                    "{variant} {key}: {name}"
                );
            }
        }
    }
}

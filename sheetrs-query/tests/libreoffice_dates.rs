//! The dates of a workbook as dates (`dates: datetime`) against `LibreOffice`.
//!
//! `sheetrs-io/tests/assets/dates*.{xlsx,ods}` hold the same serial numbers in twelve formats (a date, a date
//! and time, a time, an elapsed time, fractions of a second, the formats an XLSX numbers…) in the three
//! date systems of `LibreOffice` (1899-12-30, 1904-01-01 and, in an ODS, the 1900-01-01 of `StarCalc` 1.0).
//! `other_null_date.ods` has a null date of none of them, and `SheetRS` refuses it. Their oracles
//! (`make_oracle.py`) say what `LibreOffice` shows for each cell; the ISO 8601 text of a date value must be
//! what it shows in the formats that write one (`YYYY-MM-DD`, `HH:MM:SS`…), and the other formats must
//! give the same kind of value.

use serde_json::Value;
use sheetrs_io::Book;
use sheetrs_protocol::{DateValue, Property, Query};
use sheetrs_query::BookQuery;
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

fn assets() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../sheetrs-io/tests/assets")
}

struct Shown {
    text: String,
    format: String,
}

fn oracle(file: &str) -> BTreeMap<(u32, u32), Shown> {
    let path = assets().join("oracle").join(format!("{file}.oracle.json"));
    let json: Value = serde_json::from_slice(&std::fs::read(path).unwrap()).unwrap();
    json["cells"]
        .as_array()
        .unwrap()
        .iter()
        .map(|c| {
            (
                (
                    c["row"].as_u64().unwrap() as u32,
                    c["col"].as_u64().unwrap() as u32,
                ),
                Shown {
                    text: c["string"].as_str().unwrap().to_string(),
                    format: c["format"].as_str().unwrap().to_string(),
                },
            )
        })
        .collect()
}

fn answers(file: &str) -> BTreeMap<(u32, u32), Value> {
    let source = Book::from_file(assets().join(file)).unwrap();
    let query = Query {
        properties: vec![Property::Value],
        dates: true,
        ..Query::default()
    };
    source
        .query(&query)
        .unwrap()
        .cells
        .into_iter()
        .map(|c| ((c.row - 1, c.col - 1), c.value.unwrap()))
        .collect()
}

/// `LibreOffice` shows a year with four digits where ISO 8601 has them; anything else (`10113-09-19`,
/// `#FMT`) is a date that is not given as one.
fn iso_date(text: &str) -> bool {
    let b = text.as_bytes();
    b.len() >= 10 && b[4] == b'-' && b[7] == b'-' && text[..4].bytes().all(|c| c.is_ascii_digit())
}

fn check(file: &str) {
    let shown = oracle(file);
    let got = answers(file);
    assert!(!got.is_empty());
    let mut checked = 0;
    for (&(row, col), value) in &got {
        let libre = &shown[&(row, col)];
        let here = format!(
            "{file} row {} column {} ({}, shows {})",
            row + 1,
            col + 1,
            libre.format,
            libre.text
        );
        // An ODS stores a date before the year 1 as `-0001-12-31` (LibreOffice writes it so for the first
        // days of its range), which the reader does not take for a date: the cell is text there, as it is
        // without `dates`. It is not a number, so there is no date to give.
        if value.as_str().is_some_and(|t| t.starts_with('-')) {
            continue;
        }
        let date: Option<DateValue> = serde_json::from_value(value.clone()).ok();
        // What the same row shows in the first column (`YYYY-MM-DD`): the day, for the formats that
        // write it another way.
        let day = &shown[&(row, 0)].text;
        match libre.format.as_str() {
            "General" | "[HH]:MM:SS" => {
                assert!(value.is_number(), "{here}: a number, not {value}");
            }
            "YYYY-MM-DD" => {
                if iso_date(&libre.text) && libre.text.len() == 10 {
                    assert_eq!(date, Some(DateValue::Date(libre.text.clone())), "{here}");
                } else {
                    assert!(value.is_number(), "{here}: {value}");
                }
            }
            "YYYY-MM-DD HH:MM:SS" | "YYYY-MM-DD HH:MM:SS.000" => {
                if iso_date(&libre.text) && libre.text.find(' ') == Some(10) {
                    assert_eq!(
                        date,
                        Some(DateValue::Datetime(libre.text.replacen(' ', "T", 1))),
                        "{here}"
                    );
                } else {
                    assert!(value.is_number(), "{here}: {value}");
                }
            }
            "HH:MM:SS" | "HH:MM:SS.00" => {
                assert_eq!(date, Some(DateValue::Time(libre.text.clone())), "{here}");
            }
            "MM:SS" | "H:MM AM/PM" => {
                assert!(matches!(date, Some(DateValue::Time(_))), "{here}: {value}");
            }
            "DD/MM/YYYY" | "D-MMM-YY" | "M/D/YYYY" => {
                if iso_date(day) && day.len() == 10 {
                    assert_eq!(date, Some(DateValue::Date(day.clone())), "{here}");
                } else {
                    assert!(value.is_number(), "{here}: {value}");
                }
            }
            other => panic!("a format the fixture does not have: {other}"),
        }
        checked += 1;
    }
    // Only the cells of the reader limit above are left out.
    assert!(
        checked + 6 >= got.len(),
        "{file}: {checked} of {}",
        got.len()
    );
}

#[test]
fn dates_of_the_usual_system_are_what_libreoffice_shows() {
    check("dates.xlsx");
    check("dates.ods");
}

#[test]
fn dates_of_the_1904_system_are_what_libreoffice_shows() {
    check("dates_1904.xlsx");
    check("dates_1904.ods");
}

#[test]
fn dates_of_the_1900_01_01_system_are_what_libreoffice_shows() {
    check("dates_1900.ods");
}

/// An ODS whose null date is none of the three (`other_null_date.ods`, 1950-06-15, which
/// `LibreOffice` stores through its API) is an error that names the date, whatever is asked of it:
/// no date is numbered from a date `SheetRS` does not support.
#[test]
fn an_ods_with_another_null_date_is_an_error_that_names_it() {
    let source = Book::from_file(assets().join("other_null_date.ods")).unwrap();
    assert!(matches!(
        source.date_system(),
        Err(sheetrs_io::Error::UnsupportedDateSystem(_))
    ));
    for dates in [false, true] {
        let query = Query {
            properties: vec![Property::Value],
            dates,
            ..Query::default()
        };
        let error = source.query(&query).unwrap_err().to_string();
        assert!(error.contains("1950-06-15"), "{error}");
        assert!(error.contains("1899-12-30"), "{error}");
    }
}

#[test]
fn serial_is_the_default_and_leaves_every_value_as_it_is() {
    let source = Book::from_file(assets().join("dates.xlsx")).unwrap();
    let serial = source.query(&Query::default()).unwrap();
    assert!(
        serial
            .cells
            .iter()
            .all(|c| c.value.as_ref().unwrap().is_number())
    );
    let as_dates = answers("dates.xlsx");
    assert!(as_dates.values().any(|v| v.get("type").is_some()));
}

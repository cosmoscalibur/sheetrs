//! What `LibreOffice` writes for the same cells in an XLSX and in an ODS reads the same: the
//! formulas as Excel writes them, the dates, and the formats of the dates.
//!
//! The assets were made by `LibreOffice` 24.2.7 (the workbook built through its API and saved in
//! both formats), so the XLSX is the oracle: it holds each formula in Excel's syntax as
//! `LibreOffice` itself converts it. `formulas.*` has a sheet `F` with 41 formulas that refer to
//! other sheets in every way `LibreOffice` writes (`[T.A1]`, `[$T.A1]`, quoted names with spaces,
//! dots and apostrophes, ranges across the same or another sheet, whole columns and rows, text
//! that looks like a reference); `cells.*` has a sheet `T` with dates, times, percentages, errors
//! and formulas.

use sheetrs_io::reader::{CellValue, Source, Workbook};
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(name)
}

fn read(name: &str) -> Workbook {
    Source::from_file(asset(name))
        .unwrap()
        .read_workbook()
        .unwrap()
}

type Snapshot = BTreeMap<(u32, u32), (CellValue, Option<String>, Option<String>)>;

fn cells(workbook: &Workbook, sheet: &str) -> Snapshot {
    let sheet = workbook.sheets.iter().find(|s| s.name == sheet).unwrap();
    sheet
        .cells
        .iter()
        .map(|(at, cell)| {
            (
                at,
                (
                    cell.value.clone(),
                    cell.formula.as_deref().map(str::to_string),
                    cell.num_fmt.as_deref().map(str::to_string),
                ),
            )
        })
        .collect()
}

#[test]
fn a_formula_of_an_ods_is_written_as_the_xlsx_of_libreoffice_writes_it() {
    let (xlsx, ods) = (read("formulas.xlsx"), read("formulas.ods"));
    let (x, o) = (cells(&xlsx, "F"), cells(&ods, "F"));
    assert_eq!(x.len(), 41);
    let mut different = Vec::new();
    for (at, (_, formula, _)) in &x {
        let ods_formula = &o[at].1;
        if formula != ods_formula {
            different.push(format!("row {}: {formula:?} / {ods_formula:?}", at.0 + 1));
        }
    }
    assert!(
        different.is_empty(),
        "{} formulas differ:\n{}",
        different.len(),
        different.join("\n")
    );
}

#[test]
fn the_values_of_those_formulas_are_the_same_in_both() {
    let (xlsx, ods) = (read("formulas.xlsx"), read("formulas.ods"));
    let o = cells(&ods, "F");
    for (at, (value, _, _)) in cells(&xlsx, "F") {
        let ods_value = &o[&at].0;
        // A range in a single-cell context caches `#VALUE!`, which the XLSX reader takes for the
        // placeholder of a dynamic array (#62) and reads as empty; the ODS holds no such
        // placeholder and reads the error. It is a known difference of that rule, not of the formulas.
        if matches!(value, CellValue::Empty) && format!("{ods_value:?}") == "Error(Value)" {
            continue;
        }
        assert_eq!(
            format!("{value:?}"),
            format!("{ods_value:?}"),
            "row {}",
            at.0 + 1
        );
    }
}

/// The format of a date as the ODS gives it, and the XLSX of `LibreOffice` with its locale mark
/// (`[$-409]`, the English of the United States) taken off.
fn without_locale(format: &str) -> &str {
    format.strip_prefix("[$-409]").unwrap_or(format)
}

#[test]
fn a_date_is_the_same_number_with_the_same_format_in_both() {
    let (xlsx, ods) = (read("cells.xlsx"), read("cells.ods"));
    let (x, o) = (cells(&xlsx, "T"), cells(&ods, "T"));
    // A1 and A2 are dates and A5 and A6 are dates of the 1900 system's first days: the same
    // serial number in both, and a format that says it is a date.
    for row in [0, 1, 4, 5] {
        let (xv, _, xf) = &x[&(row, 0)];
        let (ov, _, of) = &o[&(row, 0)];
        match (xv, ov) {
            (CellValue::Number(a), CellValue::Number(b)) => {
                assert!((a - b).abs() < 1e-9, "row {}: {a} and {b}", row + 1);
            }
            other => panic!("row {}: {other:?}", row + 1),
        }
        let (xf, of) = (xf.as_deref().unwrap(), of.as_deref().unwrap());
        assert_eq!(without_locale(xf), of, "row {}", row + 1);
    }
}

/// `=SQRT(-1)` is `Err:502` and a circular reference `Err:522` in `LibreOffice`. It writes `#VALUE!`
/// for both in an XLSX, which has no such errors; an ODS says which error it is, and so does the
/// cell: the same formula reads as `LibreOffice` shows it in each file.
#[test]
fn an_error_that_only_libreoffice_has_keeps_its_code_in_an_ods() {
    let (xlsx, ods) = (read("cells.xlsx"), read("cells.ods"));
    let (x, o) = (cells(&xlsx, "T"), cells(&ods, "T"));
    for (row, code) in [(13, 502), (15, 522)] {
        assert_eq!(format!("{:?}", x[&(row, 0)].0), "Error(Value)");
        assert_eq!(
            format!("{:?}", o[&(row, 0)].0),
            format!("Error(Libre({code}))")
        );
    }
}

/// `formats.*` has 39 number formats of every family (number, grouping, percentage, scientific,
/// fraction, time, duration, date, currency, sections with colours and conditions, text) in column
/// A and the code that was asked for in column B. The ODS of `LibreOffice` and the XLSX it writes
/// give the same format code for each, locale mark aside.
#[test]
fn a_number_format_is_the_same_code_in_both() {
    let (xlsx, ods) = (read("formats.xlsx"), read("formats.ods"));
    let (x, o) = (cells(&xlsx, "T"), cells(&ods, "T"));
    let mut different = Vec::new();
    for row in 0..39 {
        let wanted = &x[&(row, 1)].0;
        let from_xlsx = x[&(row, 0)].2.as_deref().map_or("General", without_locale);
        let from_ods = o[&(row, 0)].2.as_deref().unwrap_or("General");
        // `LibreOffice` writes the format `0` of the first row as no style at all in an ODS, so the
        // cell is General there: it reads back as it is written.
        if row == 0 {
            assert_eq!((from_xlsx, from_ods), ("0", "General"));
            continue;
        }
        if from_xlsx != from_ods {
            different.push(format!("{wanted:?}: {from_xlsx:?} / {from_ods:?}"));
        }
    }
    assert!(different.is_empty(), "{}", different.join("\n"));
}

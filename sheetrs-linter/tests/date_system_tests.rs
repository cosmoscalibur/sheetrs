//! FILE1003 and the date systems on real files: the three date systems of `LibreOffice` Calc
//! (1899-12-30, 1900-01-01 and 1904-01-01) are read, and only the 1904 one is reported; an ODS with
//! any other null date is an error that names it, not a lint of dates numbered from the wrong one.
//! The files are `sheetrs-io/tests/assets/dates*.{xlsx,ods}` (made by `oracle/make_dates.py`).

use sheetrs_io::Book;
use sheetrs_linter::BookLint;
use sheetrs_linter::{Linter, RuleId};

fn assets(file: &str) -> Book {
    Book::from_file(
        concat!(env!("CARGO_MANIFEST_DIR"), "/../sheetrs-io/tests/assets/").to_owned() + file,
    )
    .unwrap()
}

fn file1003(file: &str) -> usize {
    assets(file)
        .lint(&Linter::new())
        .unwrap()
        .violations()
        .iter()
        .filter(|v| v.rule_id == RuleId::File1003)
        .count()
}

#[test]
fn only_the_1904_system_is_reported() {
    for file in ["dates_1904.xlsx", "dates_1904.ods"] {
        assert_eq!(file1003(file), 1, "{file}");
    }
    for file in ["dates.xlsx", "dates.ods", "dates_1900.ods"] {
        assert_eq!(file1003(file), 0, "{file}");
    }
}

#[test]
fn an_ods_with_another_null_date_is_an_error_that_names_the_date() {
    let error = assets("other_null_date.ods")
        .lint(&Linter::new())
        .unwrap_err()
        .to_string();
    for date in ["1950-06-15", "1899-12-30", "1900-01-01", "1904-01-01"] {
        assert!(error.contains(date), "{error}");
    }
}

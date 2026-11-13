//! Hiding and showing rows, columns and sheets, on both formats.

use sheetrs_edit::BookEdit;
use sheetrs_io::Book;
use sheetrs_io::reader::SheetState;
use std::path::{Path, PathBuf};

fn book(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("tests/assets/structure/book.{extension}"))
}

struct Seen {
    rows: Vec<(u32, u32)>,
    columns: Vec<(u32, u32)>,
}

fn seen(path: &Path, sheet: &str) -> Seen {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let sheet = workbook
        .sheets
        .into_iter()
        .find(|s| s.name == sheet)
        .unwrap();
    Seen {
        rows: sheet.hidden_rows,
        columns: sheet.hidden_columns,
    }
}

fn state(path: &Path, sheet: &str) -> SheetState {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    workbook
        .sheets
        .into_iter()
        .find(|s| s.name == sheet)
        .unwrap()
        .state
}

fn covers(ranges: &[(u32, u32)], at: u32) -> bool {
    ranges
        .iter()
        .any(|&(first, last)| (first..=last).contains(&at))
}

#[test]
fn rows_are_hidden_and_shown_in_the_file_they_are() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let hidden = dir.path().join(format!("hidden.{extension}"));
        // Row 3 exists; 700 and 701 lie beyond what the sheet declares.
        Book::from_file(book(extension))
            .unwrap()
            .hide_rows(["T!3".parse().unwrap(), "T!700:701".parse().unwrap()])
            .write(&hidden)
            .unwrap();
        let after = seen(&hidden, "T");
        for row in [2, 699, 700] {
            assert!(
                covers(&after.rows, row),
                "{extension}: row {} {:?}",
                row + 1,
                after.rows
            );
        }
        assert!(!covers(&after.rows, 3), "{extension}: {:?}", after.rows);
        assert!(!covers(&after.rows, 701), "{extension}: {:?}", after.rows);

        let shown = dir.path().join(format!("shown.{extension}"));
        Book::from_file(&hidden)
            .unwrap()
            .show_rows(["T!3".parse().unwrap(), "T!700".parse().unwrap()])
            .write(&shown)
            .unwrap();
        let after = seen(&shown, "T");
        assert!(!covers(&after.rows, 2), "{extension}: {:?}", after.rows);
        assert!(!covers(&after.rows, 699), "{extension}: {:?}", after.rows);
        assert!(covers(&after.rows, 700), "{extension}: {:?}", after.rows);
    }
}

#[test]
fn columns_are_hidden_and_shown_including_ones_the_sheet_does_not_declare() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let hidden = dir.path().join(format!("hidden.{extension}"));
        Book::from_file(book(extension))
            .unwrap()
            .hide_columns(["T!B:C".parse().unwrap(), "T!AZ".parse().unwrap()])
            .write(&hidden)
            .unwrap();
        let after = seen(&hidden, "T");
        for column in [1, 2, 51] {
            assert!(
                covers(&after.columns, column),
                "{extension}: {:?}",
                after.columns
            );
        }
        assert!(
            !covers(&after.columns, 0),
            "{extension}: {:?}",
            after.columns
        );
        assert!(
            !covers(&after.columns, 3),
            "{extension}: {:?}",
            after.columns
        );

        let shown = dir.path().join(format!("shown.{extension}"));
        Book::from_file(&hidden)
            .unwrap()
            .show_columns(["T!C".parse().unwrap()])
            .write(&shown)
            .unwrap();
        let after = seen(&shown, "T");
        assert!(
            covers(&after.columns, 1),
            "{extension}: {:?}",
            after.columns
        );
        assert!(
            !covers(&after.columns, 2),
            "{extension}: {:?}",
            after.columns
        );
    }
}

#[test]
fn sheets_are_hidden_and_shown_and_one_stays_visible() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let hidden = dir.path().join(format!("hidden.{extension}"));
        Book::from_file(book(extension))
            .unwrap()
            .hide_sheets(["T"])
            .write(&hidden)
            .unwrap();
        assert_eq!(state(&hidden, "T"), SheetState::Hidden, "{extension}");
        assert_eq!(state(&hidden, "U"), SheetState::Visible, "{extension}");

        let shown = dir.path().join(format!("shown.{extension}"));
        Book::from_file(&hidden)
            .unwrap()
            .show_sheets(["T"])
            .write(&shown)
            .unwrap();
        assert_eq!(state(&shown, "T"), SheetState::Visible, "{extension}");

        let names: Vec<String> = Book::from_file(book(extension))
            .unwrap()
            .workbook()
            .unwrap()
            .sheets
            .into_iter()
            .map(|s| s.name)
            .collect();
        let all = dir.path().join(format!("all.{extension}"));
        let error = Book::from_file(book(extension))
            .unwrap()
            .hide_sheets(names)
            .write(&all)
            .unwrap_err();
        assert!(
            error.to_string().contains("Cannot hide every sheet"),
            "{extension}: {error}"
        );
        assert!(!all.exists());
    }
}

#[test]
fn a_sheet_that_does_not_exist_is_an_error() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        for result in [
            Book::from_file(book(extension))
                .unwrap()
                .hide_sheets(["Nowhere"])
                .write(dir.path().join(format!("a.{extension}"))),
            Book::from_file(book(extension))
                .unwrap()
                .hide_rows(["Nowhere!3".parse().unwrap()])
                .write(dir.path().join(format!("b.{extension}"))),
        ] {
            assert!(result.is_err(), "{extension}");
            assert!(
                result.unwrap_err().to_string().contains("Nowhere"),
                "{extension}"
            );
        }
    }
}

#[test]
fn hiding_and_deleting_name_the_rows_of_the_input() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        // Hide row 5 and delete row 3 in one call: the hidden row is the fifth of the input, the
        // fourth of the output.
        Book::from_file(book(extension))
            .unwrap()
            .delete_rows(["T!3".parse().unwrap()])
            .hide_rows(["T!5".parse().unwrap()])
            .write(&out)
            .unwrap();
        let after = seen(&out, "T");
        assert!(covers(&after.rows, 3), "{extension}: {:?}", after.rows);
        assert!(!covers(&after.rows, 4), "{extension}: {:?}", after.rows);
    }
}

#[test]
fn the_last_change_of_a_row_wins() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        Book::from_file(book(extension))
            .unwrap()
            .hide_rows(["T!3:5".parse().unwrap()])
            .show_rows(["T!4".parse().unwrap()])
            .write(&out)
            .unwrap();
        let after = seen(&out, "T");
        assert!(
            covers(&after.rows, 2) && covers(&after.rows, 4),
            "{extension}: {:?}",
            after.rows
        );
        assert!(!covers(&after.rows, 3), "{extension}: {:?}", after.rows);
    }
}

//! The state of the rows an edit that hides or shows rows needs (`Book::row_state`): which are hidden and which hold a value or a
//! formula, in an XLSX and in an ODS, on `tests/assets/cached/visibility.*` (values in rows 1-5, 7-9 and 15 of `Data`; rows 3-4, 6
//! and 10-14 hidden, the last ones empty, which an ODS writes as a repeated row).

use sheetrs_io::Book;
use std::path::Path;

/// The rows of some bands, which an ODS gives one band per row element.
fn rows_of(bands: &[(u32, u32)]) -> Vec<u32> {
    bands
        .iter()
        .flat_map(|&(first, last)| first..=last)
        .collect()
}

fn state(extension: &str, last: u32) -> sheetrs_io::RowState {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/cached")
        .join(format!("visibility.{extension}"));
    let book = Book::from_file(path).unwrap();
    book.row_state(&[("Data", 0, last)])
        .unwrap()
        .remove("Data")
        .unwrap()
}

#[test]
fn the_hidden_rows_and_the_rows_with_a_value_are_read_in_both_formats() {
    for extension in ["xlsx", "ods"] {
        let rows = state(extension, 30);
        assert_eq!(
            rows_of(&rows.hidden),
            [2, 3, 5, 9, 10, 11, 12, 13],
            "{extension}: hidden"
        );
        assert_eq!(
            rows_of(&rows.filled),
            [0, 1, 2, 3, 4, 6, 7, 8, 14],
            "{extension}: filled"
        );
    }
}

#[test]
fn an_xlsx_is_read_up_to_the_last_row_asked_for() {
    let rows = state("xlsx", 4);
    assert_eq!(rows.hidden, vec![(2, 3)]);
    assert_eq!(rows.filled, vec![(0, 4)]);
}

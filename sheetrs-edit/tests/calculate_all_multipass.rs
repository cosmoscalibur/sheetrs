//! `calculate_all` survives the plans that run in several passes over the file.
//!
//! Each pass of a plan is a book of its own (`Book::with_calculate_all_of` copies the mode to it): a pass that lost the
//! flag would keep, for the formulas nothing reaches, the result the file caches. `cached/book.*` caches `Calc!A6` as 140
//! (`rate*2`, where `rate` is 7: 14) and `Calc!A8` as 99 (`ROW()` of row 8 is 8), values no calculation gives, so which one
//! a plan wrote says whether every pass calculated everything.

use serde_json::json;
use sheetrs_edit::{BookEdit, Calculation};
use sheetrs_io::Book;
use sheetrs_query::{BookQuery, RangeOptions};
use std::path::Path;

const EXTENSIONS: [&str; 2] = ["xlsx", "ods"];

fn written(extension: &str, plan: impl Fn(&mut Book)) -> (f64, f64) {
    let path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("tests/assets/cached/book.{extension}"));
    let mut book = Book::from_file(path).unwrap().with_calculate_all(true);
    plan(&mut book);
    let mut bytes = Vec::new();
    book.write_to(&mut bytes).unwrap();
    let blocks = Book::from_bytes(bytes)
        .unwrap()
        .range(["Calc!A6", "Calc!A8"], &RangeOptions::default())
        .unwrap();
    let at = |i: usize| {
        blocks[i].values.as_ref().unwrap()[0][0]
            .as_ref()
            .and_then(serde_json::Value::as_f64)
            .unwrap()
    };
    (at(0), at(1))
}

#[test]
fn hide_then_delete_and_convert_calculates_every_formula_in_each_pass() {
    // The visibility pass, then the pass that deletes and converts: two passes.
    for extension in EXTENSIONS {
        let got = written(extension, |book| {
            book.hide_rows(["Data!2".parse().unwrap()]);
            book.delete_columns(["Data!D".parse().unwrap()]);
            book.formula_to_value(["Calc!A:A".parse().unwrap()]);
        });
        assert_eq!(got, (14.0, 8.0), "{extension}");
    }
}

#[test]
fn calculate_hide_delete_and_convert_calculates_every_formula_in_each_pass() {
    // The calculation of the inputs is a pass of its own, the visibility another, and the deletion and the conversion
    // come after them: three passes.
    for extension in EXTENSIONS {
        let got = written(extension, |book| {
            BookEdit::calculate(
                book,
                Calculation {
                    inputs: json!({"Data!A1": 10}).as_object().cloned().unwrap(),
                    ..Calculation::default()
                },
            );
            book.hide_rows(["Data!2".parse().unwrap()]);
            book.delete_columns(["Data!D".parse().unwrap()]);
            book.formula_to_value(["Calc!A:A".parse().unwrap()]);
        });
        assert_eq!(got, (14.0, 8.0), "{extension}");
    }
}

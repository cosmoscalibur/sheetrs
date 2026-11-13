//! The operations of a `Book` are written against the workbook as it is and the planner decides what
//! runs in which order: the order they were pushed in changes nothing, down to the bytes of the
//! file written.

use sheetrs_edit::{BookEdit, Calculation};
use sheetrs_io::Book;
use sheetrs_protocol::{Find, Reference, Replace, Scope};
use std::path::{Path, PathBuf};

fn structure(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/structure")
        .join(format!("book.{extension}"))
}

/// The operations of the test, each as a way to push it: a deletion of rows, of columns, of a sheet
/// and of a name, a conversion, a replacement and a calculation.
fn push(book: &mut Book, operation: usize) {
    match operation {
        0 => book.delete_rows(["T!2".parse().unwrap()]),
        1 => book.delete_columns(["U!B".parse().unwrap()]),
        2 => book.delete_sheets(["V"]),
        3 => book.delete_names(["NameOne"]),
        4 => book.formula_to_value(["F".parse().unwrap()]),
        5 => book.replace(
            [Scope::Range("T!A3".parse::<Reference>().unwrap())],
            [Replace {
                find: Find::literal("1"),
                with: "one".into(),
            }],
        ),
        6 => book.calculate(Calculation::default()),
        _ => unreachable!(),
    };
}

fn written(extension: &str, order: &[usize]) -> Vec<u8> {
    let mut book = Book::from_file(structure(extension)).unwrap();
    for &operation in order {
        push(&mut book, operation);
    }
    let mut bytes = Vec::new();
    book.write_to(&mut bytes).unwrap();
    bytes
}

/// A deterministic shuffle, so that a failing order can be run again.
fn shuffled(mut items: Vec<usize>, mut seed: u64) -> Vec<usize> {
    for i in (1..items.len()).rev() {
        seed = seed
            .wrapping_mul(6_364_136_223_846_793_005)
            .wrapping_add(1_442_695_040_888_963_407);
        items.swap(i, ((seed >> 33) as usize) % (i + 1));
    }
    items
}

#[test]
fn permuting_the_order_the_operations_are_pushed_in_gives_byte_identical_files() {
    let operations: Vec<usize> = (0..7).collect();
    for extension in ["xlsx", "ods"] {
        let reference = written(extension, &operations);
        let mut reversed = operations.clone();
        reversed.reverse();
        assert!(
            written(extension, &reversed) == reference,
            "{extension}: reversed"
        );
        for seed in 0..12 {
            let order = shuffled(operations.clone(), seed);
            assert!(
                written(extension, &order) == reference,
                "{extension}: {order:?}"
            );
        }
    }
}

#[test]
fn splitting_an_operation_into_calls_changes_nothing() {
    for extension in ["xlsx", "ods"] {
        let mut together = Book::from_file(structure(extension)).unwrap();
        together
            .delete_rows(["T!2".parse().unwrap(), "T!4:5".parse().unwrap()])
            .delete_sheets(["U", "V"]);
        let mut apart = Book::from_file(structure(extension)).unwrap();
        apart
            .delete_sheets(["V"])
            .delete_rows(["T!4:5".parse().unwrap()])
            .delete_sheets(["U"])
            .delete_rows(["T!2".parse().unwrap()]);
        let (mut a, mut b) = (Vec::new(), Vec::new());
        together.write_to(&mut a).unwrap();
        apart.write_to(&mut b).unwrap();
        assert!(a == b, "{extension}");
    }
}

#[test]
fn a_replacement_a_conversion_and_a_deletion_in_one_run_are_calculated_and_not_marked() {
    // The replacement leaves results stale, the conversion has them calculated again and the
    // deletion after it leaves some stale: the engine calculates them as the last pass, so the
    // workbook has fresh results and no mark (the mark is only for a workbook the engine cannot
    // calculate).
    let bytes = written("xlsx", &[0, 4, 5]);
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut xml = String::new();
    std::io::Read::read_to_string(&mut archive.by_name("xl/workbook.xml").unwrap(), &mut xml)
        .unwrap();
    assert!(!xml.contains("fullCalcOnLoad"), "{xml}");
    assert!(xml.matches("<calcPr").count() <= 1, "{xml}");
}

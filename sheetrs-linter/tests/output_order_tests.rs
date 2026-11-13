use sheetrs_io::Book;
use sheetrs_linter::{FormatContext, Linter};

const TEST_XLSX: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/lint_book.xlsx");
const TEST_ODS: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/lint_book.ods");

/// Lints a freshly parsed workbook, so every run gets new `HashMap` seeds.
fn lint_output(path: &str) -> Vec<String> {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let ctx = FormatContext {
        workbook: &workbook,
    };
    Linter::new()
        .lint_workbook(&workbook)
        .iter()
        .map(|v| format!("{:?} {} {}", v.scope, v.rule_id, v.format_message(&ctx)))
        .collect()
}

fn assert_stable_output(path: &str) {
    let first = lint_output(path);
    assert!(!first.is_empty());
    for _ in 0..20 {
        assert_eq!(lint_output(path), first);
    }
}

#[test]
fn test_output_order_is_stable_xlsx() {
    assert_stable_output(TEST_XLSX);
}

#[test]
fn test_output_order_is_stable_ods() {
    assert_stable_output(TEST_ODS);
}

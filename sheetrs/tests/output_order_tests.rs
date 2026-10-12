use sheetrs::reader::read_workbook_from_reader;
use sheetrs::{FormatContext, Linter};
use std::io::Cursor;

const TEST_XLSX: &[u8] = include_bytes!("../../tests/minimal_test.xlsx");
const TEST_ODS: &[u8] = include_bytes!("../../tests/minimal_test.ods");

/// Lints a freshly parsed workbook, so every run gets new `HashMap` seeds.
fn lint_output(bytes: &[u8], extension: &str) -> Vec<String> {
    let workbook = read_workbook_from_reader(Cursor::new(bytes), Some(extension)).unwrap();
    let ctx = FormatContext {
        workbook: &workbook,
    };
    Linter::new()
        .lint_workbook(&workbook)
        .unwrap()
        .iter()
        .map(|v| format!("{:?} {} {}", v.scope, v.rule_id, v.format_message(&ctx)))
        .collect()
}

fn assert_stable_output(bytes: &[u8], extension: &str) {
    let first = lint_output(bytes, extension);
    assert!(!first.is_empty());
    for _ in 0..20 {
        assert_eq!(lint_output(bytes, extension), first);
    }
}

#[test]
fn test_output_order_is_stable_xlsx() {
    assert_stable_output(TEST_XLSX, "xlsx");
}

#[test]
fn test_output_order_is_stable_ods() {
    assert_stable_output(TEST_ODS, "ods");
}

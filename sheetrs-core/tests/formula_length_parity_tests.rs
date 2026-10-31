//! The length CPX509 measures is that of the formula the user sees, the same in XLSX and ODS: the
//! namespace prefix of a function (`_xlfn.`, `COM.MICROSOFT.`) is not part of it.

use sheetrs_core::reader::read_workbook;
use sheetrs_formula::calc_print::visible_length;

/// The formula of `Sheet7!G13` of `tests/minimal_test.<extension>`.
fn formula(extension: &str) -> String {
    let path = format!(
        "{}/../tests/minimal_test.{extension}",
        env!("CARGO_MANIFEST_DIR")
    );
    let workbook = read_workbook(path).unwrap();
    let sheet = workbook.sheets.iter().find(|s| s.name == "Sheet7").unwrap();
    sheet
        .cells
        .get(&(12, 6))
        .and_then(|c| c.formula.as_deref())
        .unwrap()
        .to_string()
}

#[test]
fn a_formula_has_the_same_visible_length_in_both_formats() {
    let (xlsx, ods) = (formula("xlsx"), formula("ods"));
    assert!(xlsx.contains("_xlfn.CONCAT") && ods.contains("COM.MICROSOFT.CONCAT"));
    assert_ne!(xlsx.chars().count(), ods.chars().count());
    assert_eq!(visible_length(&xlsx), visible_length(&ods));
}

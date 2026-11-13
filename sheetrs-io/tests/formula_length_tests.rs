//! The length CPX509 measures is that of the formula the user sees: the namespace prefix of a
//! function (`_xlfn.`, `COM.MICROSOFT.`) is not part of it. What the user sees is what
//! `LibreOffice` shows for the cell of each file (`tests/assets/oracle`), which is the length of the
//! formula without its `=`.

use sheetrs_formula::calc_print::visible_length;
use sheetrs_io::Book;

/// The formula of `Complexity!B7` of `tests/lint_book.<extension>`.
fn formula(extension: &str) -> String {
    let path = format!(
        "{}/../tests/lint_book.{extension}",
        env!("CARGO_MANIFEST_DIR")
    );
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Complexity")
        .unwrap();
    sheet
        .cells
        .get(&(6, 1))
        .and_then(|c| c.formula.as_deref())
        .unwrap()
        .to_string()
}

/// The length of the formula of `Complexity!B7` as `LibreOffice` shows it.
fn libre_length(extension: &str) -> usize {
    let path = format!(
        "{}/tests/assets/oracle/lint_book.{extension}.oracle.json",
        env!("CARGO_MANIFEST_DIR")
    );
    let oracle: serde_json::Value = serde_json::from_slice(&std::fs::read(path).unwrap()).unwrap();
    let cell = oracle["cells"]
        .as_array()
        .unwrap()
        .iter()
        .find(|c| c["sheet"] == "Complexity" && c["row"] == 6 && c["col"] == 1)
        .unwrap();
    cell["formula"]
        .as_str()
        .unwrap()
        .trim_start_matches('=')
        .chars()
        .count()
}

#[test]
fn the_visible_length_of_a_formula_is_the_one_libreoffice_shows() {
    let (xlsx, ods) = (formula("xlsx"), formula("ods"));
    // The prefixes the files write are not in what the user sees.
    assert!(xlsx.contains("_xlfn.CONCAT") && ods.contains("COM.MICROSOFT.CONCAT"));
    for (extension, formula) in [("xlsx", xlsx), ("ods", ods)] {
        assert!(
            formula.chars().count() > libre_length(extension),
            "{extension}"
        );
        assert_eq!(
            visible_length(&formula),
            libre_length(extension),
            "{extension}"
        );
    }
}

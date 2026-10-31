//! The document properties that the readers give, named the same in XLSX and ODS.

use sheetrs_core::reader::read_workbook;

fn properties(extension: &str) -> Vec<(String, String)> {
    let path = format!(
        "{}/../tests/minimal_test.{extension}",
        env!("CARGO_MANIFEST_DIR")
    );
    read_workbook(path).unwrap().properties
}

#[test]
fn the_properties_of_both_formats_share_their_names() {
    let (xlsx, ods) = (properties("xlsx"), properties("ods"));
    for key in ["created", "modified", "generator"] {
        assert!(xlsx.iter().any(|(k, _)| k == key), "xlsx {key}");
        assert!(ods.iter().any(|(k, _)| k == key), "ods {key}");
    }
}

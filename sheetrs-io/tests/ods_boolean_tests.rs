//! The `office:boolean-value` of an ODS cell: `true`, `false`, and the number that `LibreOffice`
//! writes when the number of a cell has a boolean format (`IF({0;1},TRUE(),5)`: the cell shows TRUE
//! and its attribute is `5`). 0 and 1 are the booleans; any other number is that number, as when
//! `LibreOffice` reads it (its value is 5, its format shows TRUE).

use std::fmt::Write as _;
use std::io::{Cursor, Write};

use sheetrs_io::Book;
use sheetrs_io::reader::{Cell, CellSink, CellValue};
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

#[derive(Default)]
struct Collect {
    cells: Vec<((u32, u32), Cell)>,
}

impl CellSink for Collect {
    fn cell(&mut self, _sheet: u16, at: (u32, u32), cell: Cell) {
        self.cells.push((at, cell));
    }
}

fn ods(booleans: &[&str]) -> Book {
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let mut part = |name: &str, text: String| {
        zip.start_file(name, SimpleFileOptions::default()).unwrap();
        zip.write_all(text.as_bytes()).unwrap();
    };
    part(
        "mimetype",
        "application/vnd.oasis.opendocument.spreadsheet".to_string(),
    );
    part(
        "META-INF/manifest.xml",
        r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#.to_string(),
    );
    let mut rows = String::new();
    for b in booleans {
        write!(
            rows,
            r#"<table:table-row><table:table-cell office:value-type="boolean" office:boolean-value="{b}"><text:p>x</text:p></table:table-cell></table:table-row>"#
        )
        .unwrap();
    }
    part(
        "content.xml",
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="B">{rows}</table:table></office:spreadsheet></office:body></office:document-content>"#
        ),
    );
    Book::from_bytes(zip.finish().unwrap().into_inner()).unwrap()
}

#[test]
fn a_boolean_value_is_true_or_false_and_a_number_that_is_not_0_or_1_is_that_number() {
    let cases = [
        ("true", CellValue::Boolean(true)),
        ("false", CellValue::Boolean(false)),
        ("1", CellValue::Boolean(true)),
        ("0", CellValue::Boolean(false)),
        ("0.0", CellValue::Boolean(false)),
        // What `LibreOffice` writes for a number that has a boolean format: its value is that number.
        ("5", CellValue::Number(5.0)),
        ("-1", CellValue::Number(-1.0)),
        ("0.5", CellValue::Number(0.5)),
        ("43831", CellValue::Number(43831.0)),
    ];
    let source = ods(&cases.clone().map(|(text, _)| text));
    let mut sink = Collect::default();
    source.workbook_into(&mut sink).unwrap();
    assert_eq!(sink.cells.len(), cases.len());
    for ((_, cell), (text, expected)) in sink.cells.iter().zip(cases) {
        assert_eq!(cell.value, expected, "office:boolean-value=\"{text}\"");
    }
}

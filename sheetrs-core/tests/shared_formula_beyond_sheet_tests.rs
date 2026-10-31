//! A shared formula whose cell would read a reference outside the sheet: LibreOffice computes
//! `#REF!` there. The reader writes `Sheet!#REF!` in the formula, and the linter reports the error
//! cell.

use sheetrs_core::reader::read_workbook_from_reader;
use sheetrs_core::{Linter, RuleId};
use std::io::{Cursor, Write};

/// An XLSX with one sheet `S` holding `sheet_data`.
fn xlsx_with(sheet_data: &str) -> Cursor<Vec<u8>> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_string()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.to_string()),
        ("xl/workbook.xml", r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_string()),
        ("xl/_rels/workbook.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#.to_string()),
        ("xl/worksheets/sheet1.xml", format!(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{sheet_data}</sheetData></worksheet>"#)),
    ];
    for (name, body) in parts {
        zip.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    Cursor::new(zip.finish().unwrap().into_inner())
}

#[test]
fn the_reader_writes_a_ref_error_and_the_linter_reports_the_cell() {
    // `=XFD2` in XFC1 shared with XFD1, whose copy would read column XFE.
    let book = read_workbook_from_reader(
        xlsx_with(
            r#"<row r="1"><c r="XFC1"><f t="shared" ref="XFC1:XFD1" si="0">XFD2</f><v>0</v></c><c r="XFD1"><f t="shared" si="0"/><v>0</v></c></row>"#,
        ),
        Some("xlsx"),
    )
    .unwrap();
    let sheet = &book.sheets[0];
    assert_eq!(
        sheet.get_cell(0, 16_382).unwrap().as_formula(),
        Some("XFD2")
    );
    assert_eq!(
        sheet.get_cell(0, 16_383).unwrap().as_formula(),
        Some("S!#REF!")
    );
    let violations = Linter::new().lint_workbook(&book).unwrap();
    assert!(
        violations.iter().any(|v| v.rule_id == RuleId::Err102),
        "{violations:?}"
    );
}

//! The date system a workbook states: `date1904` on the `workbookPr` of an XLSX and the
//! `table:null-date` of an ODS. The books are made here, in memory, as small as the readers allow:
//! only the setting matters, so no workbook is kept for it.

use std::io::{Cursor, Write};

use sheetrs_io::Book;
use sheetrs_io::{DateSystem, Error};
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

fn archive(parts: &[(&str, String)]) -> Vec<u8> {
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, text) in parts {
        zip.start_file(*name, SimpleFileOptions::default()).unwrap();
        zip.write_all(text.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// An XLSX with one sheet; `pr` is the whole `workbookPr` element (or nothing).
fn xlsx(pr: &str) -> Book {
    let bytes = archive(&[
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_string(),
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.to_string(),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{pr}<sheets><sheet name="Dates" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#.to_string(),
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>41000</v></c></row></sheetData></worksheet>"#.to_string(),
        ),
    ]);
    Book::from_bytes(bytes).unwrap()
}

/// An ODS with one table; `null_date` is the `table:null-date` element (or nothing).
fn ods(null_date: &str) -> Book {
    let bytes = archive(&[
        (
            "mimetype",
            "application/vnd.oasis.opendocument.spreadsheet".to_string(),
        ),
        (
            "META-INF/manifest.xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#.to_string(),
        ),
        (
            "content.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3"><office:body><office:spreadsheet><table:calculation-settings>{null_date}</table:calculation-settings><table:table table:name="Dates"><table:table-row><table:table-cell office:value-type="float" office:value="41000"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
            ),
        ),
    ]);
    Book::from_bytes(bytes).unwrap()
}

/// The date system from `Book::date_system` and from the workbook read: the two agree.
fn systems(source: &Book) -> (DateSystem, DateSystem) {
    (
        source.date_system().unwrap(),
        source.workbook().unwrap().date_system,
    )
}

fn null_date(date: &str) -> String {
    format!(r#"<table:null-date table:value-type="date" table:date-value="{date}"/>"#)
}

#[test]
fn an_xlsx_states_1904_with_date1904_on_the_workbook_properties() {
    for (pr, expected) in [
        (r#"<workbookPr date1904="1"/>"#, DateSystem::Null1904),
        (r#"<workbookPr date1904="true"/>"#, DateSystem::Null1904),
        (r#"<workbookPr date1904="0"/>"#, DateSystem::Null1899),
        (r#"<workbookPr date1904="false"/>"#, DateSystem::Null1899),
        (
            r#"<workbookPr defaultThemeVersion="124226"/>"#,
            DateSystem::Null1899,
        ),
        ("", DateSystem::Null1899),
    ] {
        assert_eq!(systems(&xlsx(pr)), (expected, expected), "{pr}");
    }
}

#[test]
fn an_ods_states_one_of_the_three_date_systems_in_its_null_date() {
    for (date, expected) in [
        // The default of LibreOffice, StarCalc 1.0 and the 1904 system.
        ("1899-12-30", DateSystem::Null1899),
        ("1900-01-01", DateSystem::Null1900),
        ("1904-01-01", DateSystem::Null1904),
    ] {
        assert_eq!(
            systems(&ods(&null_date(date))),
            (expected, expected),
            "{date}"
        );
    }
    // Without the element the default is the 1899 system.
    assert_eq!(
        systems(&ods("")),
        (DateSystem::Null1899, DateSystem::Null1899)
    );
}

/// Any other null date is an error that names it, from every way of asking the workbook, before a
/// date is numbered from it: `LibreOffice` can store one (its API, a crafted file) and `SheetRS` does
/// not support it (`docs/approved-differences.md`).
#[test]
fn an_ods_with_another_null_date_is_an_error_that_names_it() {
    for date in [
        "1950-06-15",
        "2000-01-01",
        "1899-12-31",
        "1904-01-02",
        "1901-01-01",
    ] {
        let source = ods(&null_date(date));
        for error in [
            source.date_system().unwrap_err(),
            source.workbook().unwrap_err(),
        ] {
            assert!(
                matches!(error, Error::UnsupportedDateSystem(_)),
                "{date}: {error:?}"
            );
            let text = error.to_string();
            assert!(text.contains(date), "{text}");
            for supported in ["1899-12-30", "1900-01-01", "1904-01-01"] {
                assert!(text.contains(supported), "{text}");
            }
        }
    }
}

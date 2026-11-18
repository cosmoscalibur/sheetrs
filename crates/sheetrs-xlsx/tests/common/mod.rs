use std::error::Error as StdError;
use std::io::{Cursor, Write};

use zip::CompressionMethod;
use zip::write::{SimpleFileOptions, ZipWriter};

pub const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

pub fn package(parts: &[(&str, String)]) -> Result<Cursor<Vec<u8>>, Box<dyn StdError>> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    for (name, content) in parts {
        writer.start_file(*name, options)?;
        writer.write_all(content.as_bytes())?;
    }
    Ok(writer.finish()?)
}

pub fn workbook_xml(sheets: &str, properties: &str) -> String {
    format!(
        r#"<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">{properties}<sheets>{sheets}</sheets></workbook>"#
    )
}

pub fn rels_xml(relationships: &str) -> String {
    format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationships}</Relationships>"#
    )
}

pub fn worksheet_xml(rows: &str) -> String {
    format!(r#"<worksheet xmlns="{MAIN_NS}"><sheetData>{rows}</sheetData></worksheet>"#)
}

/// A one-sheet package with an optional shared strings part and extra children for the
/// `<workbook>` element (such as `<workbookPr date1904="1"/>`).
pub fn single_sheet_package_with(
    rows: &str,
    shared_strings: Option<&str>,
    workbook_properties: &str,
) -> Result<Cursor<Vec<u8>>, Box<dyn StdError>> {
    let mut workbook_relationships = format!(
        r#"<Relationship Id="rId1" Type="{REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>"#
    );
    let mut content_types = String::from(
        r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
    );
    let mut parts = vec![
        (
            "xl/workbook.xml",
            workbook_xml(
                r#"<sheet name="S1" sheetId="1" r:id="rId1"/>"#,
                workbook_properties,
            ),
        ),
        ("xl/worksheets/sheet1.xml", worksheet_xml(rows)),
    ];
    if let Some(strings) = shared_strings {
        workbook_relationships = format!(
            r#"{workbook_relationships}<Relationship Id="rId2" Type="{REL_NS}/sharedStrings" Target="sharedStrings.xml"/>"#
        );
        content_types.push_str(
            r#"<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"#,
        );
        parts.push((
            "xl/sharedStrings.xml",
            format!(r#"<sst xmlns="{MAIN_NS}">{strings}</sst>"#),
        ));
    }
    parts.push((
        "xl/_rels/workbook.xml.rels",
        rels_xml(&workbook_relationships),
    ));
    // The two package-level parts make the file loadable by other consumers.
    parts.push((
        "[Content_Types].xml",
        format!(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{content_types}</Types>"#
        ),
    ));
    parts.push((
        "_rels/.rels",
        rels_xml(&format!(
            r#"<Relationship Id="rId1" Type="{REL_NS}/officeDocument" Target="xl/workbook.xml"/>"#
        )),
    ));
    package(&parts)
}

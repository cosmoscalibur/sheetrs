//! A workbook written after it was read keeps what the model does not interpret.

use std::collections::BTreeMap;
use std::error::Error as StdError;
use std::io::{Cursor, Read, Write};

use sheetrs_core::{Cell, CellPos, Value};
use zip::write::{SimpleFileOptions, ZipWriter};
use zip::{CompressionMethod, ZipArchive};

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3"><office:automatic-styles><style:style style:name="ce1" style:family="table-cell"/></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="S1"><table:table-column table:number-columns-repeated="3"/><table:table-row><table:table-cell table:style-name="ce1" office:value-type="float" office:value="1" table:number-columns-spanned="2"><text:p>1</text:p></table:table-cell><table:covered-table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;

fn archive() -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    writer.start_file("mimetype", stored)?;
    writer.write_all(b"application/vnd.oasis.opendocument.spreadsheet")?;
    writer.start_file("content.xml", deflated)?;
    writer.write_all(CONTENT.as_bytes())?;
    writer.start_file("styles.xml", deflated)?;
    writer.write_all(b"<office:document-styles/>")?;
    writer.start_file("Pictures/logo.bin", deflated)?;
    writer.write_all(&[0, 159, 146, 150, 255])?;
    writer.start_file("META-INF/manifest.xml", deflated)?;
    writer.write_all(b"<manifest:manifest/>")?;
    Ok(writer.finish()?.into_inner())
}

fn entries(bytes: Vec<u8>) -> Result<BTreeMap<String, Vec<u8>>, Box<dyn StdError>> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut found = BTreeMap::new();
    for index in 0..archive.len() {
        let mut file = archive.by_index(index)?;
        let mut content = Vec::new();
        file.read_to_end(&mut content)?;
        found.insert(file.name()?.into_owned(), content);
    }
    Ok(found)
}

#[test]
fn an_unchanged_workbook_is_written_as_it_was_read_over_ods() -> Result<(), Box<dyn StdError>> {
    let workbook = sheetrs_ods::read(Cursor::new(archive()?))?;
    let mut written = Cursor::new(Vec::new());
    sheetrs_ods::write(&workbook, &mut written)?;
    let written = written.into_inner();
    let (before, after) = (entries(archive()?)?, entries(written.clone())?);
    assert_eq!(before, after);
    let first = ZipArchive::new(Cursor::new(written))?
        .by_index(0)?
        .name()?
        .into_owned();
    assert_eq!(first, "mimetype");
    Ok(())
}

#[test]
fn a_changed_workbook_keeps_the_other_parts_over_ods() -> Result<(), Box<dyn StdError>> {
    let mut workbook = sheetrs_ods::read(Cursor::new(archive()?))?;
    let sheet = workbook.sheet("S1").ok_or("sheet")?;
    let mut replacement = sheetrs_core::Workbook::new();
    let mut builder = replacement.sheet_builder("S1");
    builder.absorb(&sheet)?;
    builder.set(
        CellPos { row: 4, col: 0 },
        &Cell {
            value: Value::Number(9.0),
            formula: None,
            stale: false,
        },
    )?;
    builder.finish();
    // The package is read again with the new cell in a workbook that has the same package.
    let package = workbook.opaque().cloned().ok_or("package")?;
    replacement.set_opaque(package);
    workbook = replacement;
    let mut written = Cursor::new(Vec::new());
    sheetrs_ods::write(&workbook, &mut written)?;
    let (before, after) = (entries(archive()?)?, entries(written.into_inner())?);
    for name in ["styles.xml", "Pictures/logo.bin", "META-INF/manifest.xml"] {
        assert_eq!(before.get(name), after.get(name), "{name}");
    }
    assert_ne!(before.get("content.xml"), after.get("content.xml"));
    Ok(())
}

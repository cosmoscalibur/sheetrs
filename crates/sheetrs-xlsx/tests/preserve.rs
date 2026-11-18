//! A workbook written after it was read keeps what the model does not interpret.

use std::collections::BTreeMap;
use std::error::Error as StdError;
use std::io::{Cursor, Read, Write};

use sheetrs_core::{CellPos, Value};
use zip::write::{SimpleFileOptions, ZipWriter};
use zip::{CompressionMethod, ZipArchive};

fn package(parts: &[(&str, String)]) -> Result<Cursor<Vec<u8>>, Box<dyn StdError>> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    for (name, content) in parts {
        writer.start_file(*name, options)?;
        writer.write_all(content.as_bytes())?;
    }
    Ok(writer.finish()?)
}

const HEAD: &str = r#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><sheetPr><tabColor rgb="FFFF0000"/></sheetPr><dimension ref="A1:B2"/><cols><col min="1" max="1" width="20" customWidth="1"/></cols>"#;
const TAIL: &str = r#"<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><drawing r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></worksheet>"#;
const ROWS: &str = r#"<row r="1" spans="1:2" ht="30" customHeight="1" x14ac:dyDescent="0.25"><c r="A1" s="3" t="s"><v>0</v></c><c r="B1" s="4"><v>2</v></c></row><row r="2" hidden="1"/><row r="3"><c r="A3" s="5"/></row>"#;

fn archive() -> Result<Cursor<Vec<u8>>, Box<dyn StdError>> {
    package(&[
        ("[Content_Types].xml", "<Types/>".to_owned()),
        ("_rels/.rels", "<Relationships/>".to_owned()),
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr/><definedNames><definedName name="N">S1!$A$1</definedName></definedNames><sheets><sheet name="S1" sheetId="7" r:id="rId1"/></sheets><calcPr calcId="1"/></workbook>"#.to_owned(),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#.to_owned(),
        ),
        ("xl/styles.xml", "<styleSheet/>".to_owned()),
        ("xl/drawings/drawing1.xml", "<drawing>\u{1F600}</drawing>".to_owned()),
        ("xl/sharedStrings.xml", r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>hola</t></si></sst>"#.to_owned()),
        ("xl/worksheets/sheet1.xml", format!("{HEAD}<sheetData>{ROWS}</sheetData>{TAIL}")),
    ])
}

fn entries(bytes: Vec<u8>) -> Result<BTreeMap<String, String>, Box<dyn StdError>> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut found = BTreeMap::new();
    for index in 0..archive.len() {
        let mut file = archive.by_index(index)?;
        let mut text = String::new();
        file.read_to_string(&mut text)?;
        found.insert(file.name()?.into_owned(), text);
    }
    Ok(found)
}

#[test]
fn what_the_model_does_not_hold_is_written_back_over_xlsx() -> Result<(), Box<dyn StdError>> {
    let workbook = sheetrs_xlsx::read(archive()?)?;
    let mut written = Cursor::new(Vec::new());
    sheetrs_xlsx::write(&workbook, &mut written)?;
    let parts = entries(written.into_inner())?;
    let original = entries(archive()?.into_inner())?;
    for name in [
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/_rels/workbook.xml.rels",
        "xl/styles.xml",
        "xl/drawings/drawing1.xml",
    ] {
        assert_eq!(parts.get(name), original.get(name), "{name}");
    }
    let workbook_part = parts.get("xl/workbook.xml").ok_or("workbook part")?;
    assert!(workbook_part.contains(r#"<definedName name="N">S1!$A$1</definedName>"#));
    assert!(workbook_part.contains(r#"<sheet name="S1" sheetId="7" r:id="rId1"/>"#));
    assert!(workbook_part.contains(r#"<calcPr calcId="1"/>"#));
    let sheet = parts.get("xl/worksheets/sheet1.xml").ok_or("sheet part")?;
    assert!(sheet.starts_with(&HEAD.replace("A1:B2", "A1:B3")));
    assert!(sheet.ends_with(TAIL));
    assert!(sheet.contains(r#"<row r="1" ht="30" customHeight="1" x14ac:dyDescent="0.25">"#));
    assert!(sheet.contains(r#"<row r="2" hidden="1"></row>"#));
    assert!(sheet.contains(r#"<c r="A1" s="3" t="s"><v>0</v></c><c r="B1" s="4"><v>2</v></c>"#));
    assert!(sheet.contains(r#"<c r="A3" s="5"/>"#));
    Ok(())
}

#[test]
fn a_changed_cell_keeps_its_style_over_xlsx() -> Result<(), Box<dyn StdError>> {
    let workbook = sheetrs_xlsx::read(archive()?)?;
    let sheet = workbook.sheet("S1").ok_or("sheet")?;
    assert_eq!(
        sheet
            .get(CellPos { row: 0, col: 1 })
            .map(|cell| cell.to_cell().value),
        Some(Value::Number(2.0))
    );
    let mut written = Cursor::new(Vec::new());
    sheetrs_xlsx::write(&workbook, &mut written)?;
    let workbook = sheetrs_xlsx::read(Cursor::new(written.into_inner()))?;
    let again = workbook.sheet("S1").ok_or("sheet")?;
    assert_eq!(again.styled_cells().count(), 3);
    Ok(())
}

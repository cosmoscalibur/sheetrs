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

const RICH: &str = "<si><r><rPr><b/></rPr><t>ho</t></r><r><t>la</t></r></si>";

#[test]
fn formatted_text_keeps_its_formatting_over_xlsx() -> Result<(), Box<dyn StdError>> {
    let source = package(&[
        ("[Content_Types].xml", "<Types/>".to_owned()),
        ("_rels/.rels", "<Relationships/>".to_owned()),
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_owned(),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#.to_owned(),
        ),
        (
            "xl/sharedStrings.xml",
            format!(r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{RICH}<si><t>hola</t></si></sst>"#),
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c></row></sheetData></worksheet>"#.to_owned(),
        ),
    ])?;
    let workbook = sheetrs_xlsx::read(source)?;
    let sheet = workbook.sheet("S1").ok_or("sheet")?;
    let text = |row| {
        sheet
            .get(CellPos { row, col: 0 })
            .map(|cell| cell.to_cell().value)
    };
    assert_eq!(text(0), Some(Value::Text("hola".into())));
    assert_eq!(text(1), Some(Value::Text("hola".into())));
    let mut written = Cursor::new(Vec::new());
    sheetrs_xlsx::write(&workbook, &mut written)?;
    let parts = entries(written.into_inner())?;
    let strings = parts.get("xl/sharedStrings.xml").ok_or("strings")?;
    assert!(strings.contains(RICH), "{strings}");
    assert!(
        strings.contains(r#"<si><t xml:space="preserve">hola</t></si>"#),
        "{strings}"
    );
    Ok(())
}

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn small_package(
    workbook: &str,
    sheet: &str,
    extra: &[(&str, String)],
) -> Result<Cursor<Vec<u8>>, Box<dyn StdError>> {
    let mut parts = vec![
        ("[Content_Types].xml", "<Types/>".to_owned()),
        ("_rels/.rels", "<Relationships/>".to_owned()),
        ("xl/workbook.xml", workbook.to_owned()),
        (
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
            ),
        ),
        ("xl/worksheets/sheet1.xml", sheet.to_owned()),
    ];
    parts.extend(extra.iter().map(|(name, text)| (*name, text.clone())));
    package(&parts)
}

fn plain_workbook() -> String {
    format!(
        r#"<workbook xmlns="{MAIN}" xmlns:r="{RELS}"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>"#
    )
}

fn written(
    workbook: &sheetrs_core::Workbook,
) -> Result<BTreeMap<String, String>, Box<dyn StdError>> {
    let mut bytes = Cursor::new(Vec::new());
    sheetrs_xlsx::write(workbook, &mut bytes)?;
    entries(bytes.into_inner())
}

#[test]
fn only_the_first_cell_of_an_array_is_written_as_a_formula_over_xlsx()
-> Result<(), Box<dyn StdError>> {
    let sheet = format!(
        r#"<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><f t="array" ref="A1:B2">1+1</f><v>2</v></c></row></sheetData></worksheet>"#
    );
    let workbook = sheetrs_xlsx::read(small_package(&plain_workbook(), &sheet, &[])?)?;
    let parts = written(&workbook)?;
    let text = parts.get("xl/worksheets/sheet1.xml").ok_or("sheet")?;
    assert_eq!(text.matches("<f").count(), 1, "{text}");
    assert!(
        text.contains(r#"<f t="array" ref="A1:B2">1+1</f>"#),
        "{text}"
    );
    Ok(())
}

#[test]
fn a_sheet_part_without_cells_is_written_as_one_document_over_xlsx() -> Result<(), Box<dyn StdError>>
{
    let sheet = format!(r#"<worksheet xmlns="{MAIN}"><dimension ref="A1"/></worksheet>"#);
    let workbook = sheetrs_xlsx::read(small_package(&plain_workbook(), &sheet, &[])?)?;
    let parts = written(&workbook)?;
    let text = parts.get("xl/worksheets/sheet1.xml").ok_or("sheet")?;
    assert_eq!(text.matches("<worksheet").count(), 1, "{text}");
    sheetrs_xlsx::read(Cursor::new({
        let mut bytes = Cursor::new(Vec::new());
        sheetrs_xlsx::write(&workbook, &mut bytes)?;
        bytes.into_inner()
    }))?;
    Ok(())
}

#[test]
fn names_in_another_grammar_are_refused_over_xlsx() {
    let mut workbook = sheetrs_core::Workbook::new();
    workbook.set_formula_grammar(sheetrs_core::FormulaGrammar::OpenFormula);
    workbook.sheet_builder("S1").finish();
    workbook.define_name_replacing(sheetrs_core::NamedExpression {
        name: "N".to_owned(),
        scope: None,
        expression: "$S1.$A$1".to_owned(),
        hidden: false,
    });
    let mut bytes = Cursor::new(Vec::new());
    assert!(sheetrs_xlsx::write(&workbook, &mut bytes).is_err());
}

#[test]
fn a_package_with_prefixed_names_is_written_from_the_model_over_xlsx()
-> Result<(), Box<dyn StdError>> {
    let workbook_part = format!(
        r#"<x:workbook xmlns:x="{MAIN}" xmlns:r="{RELS}"><x:sheets><x:sheet name="S1" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>"#
    );
    let sheet = format!(
        r#"<x:worksheet xmlns:x="{MAIN}"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>3</x:v></x:c></x:row></x:sheetData></x:worksheet>"#
    );
    let workbook = sheetrs_xlsx::read(small_package(&workbook_part, &sheet, &[])?)?;
    let parts = written(&workbook)?;
    let again = {
        let mut bytes = Cursor::new(Vec::new());
        sheetrs_xlsx::write(&workbook, &mut bytes)?;
        sheetrs_xlsx::read(Cursor::new(bytes.into_inner()))?
    };
    let value = again
        .sheet("S1")
        .and_then(|sheet| sheet.get(CellPos { row: 0, col: 0 }))
        .map(|cell| cell.to_cell().value);
    assert_eq!(value, Some(Value::Number(3.0)));
    assert!(
        !parts
            .get("xl/workbook.xml")
            .ok_or("workbook")?
            .contains("x:")
    );
    Ok(())
}

#[test]
fn the_calculation_chain_is_dropped_whatever_its_spelling_over_xlsx()
-> Result<(), Box<dyn StdError>> {
    let sheet = format!(r#"<worksheet xmlns="{MAIN}"><sheetData/></worksheet>"#);
    let types = "<Types><Override\n PartName=\"/xl/calcChain.xml\" ContentType=\"x\"/><Default Extension=\"xml\" ContentType=\"y\"/></Types>";
    let mut parts = vec![
        ("[Content_Types].xml", types.to_owned()),
        ("_rels/.rels", "<Relationships/>".to_owned()),
        ("xl/workbook.xml", plain_workbook()),
        (
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id='rId9' Type='{RELS}/calcChain' Target='calcChain.xml'/></Relationships>"#
            ),
        ),
        ("xl/worksheets/sheet1.xml", sheet),
        ("xl/calcChain.xml", "<calcChain/>".to_owned()),
    ];
    let workbook = sheetrs_xlsx::read(package(&parts)?)?;
    parts.clear();
    let written = written(&workbook)?;
    assert!(!written.contains_key("xl/calcChain.xml"));
    for text in written.values() {
        assert!(!text.contains("calcChain"), "{text}");
    }
    Ok(())
}

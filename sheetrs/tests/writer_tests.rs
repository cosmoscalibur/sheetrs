use sheetrs::reader::{CellValue, read_workbook};
use sheetrs::writer::{WorkbookModifications, modify_workbook};
use std::collections::HashSet;
use std::fmt::Write as _;
use std::fs::File;
use std::io::Write;
use std::path::Path;
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

// Helper to create a minimal valid XLSX file for testing
fn create_mock_xlsx(path: &Path, sheets: &[&str], ranges: &[(&str, &str)]) -> anyhow::Result<()> {
    let file = File::create(path)?;
    let mut zip = ZipWriter::new(file);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);

    // 1. [Content_Types].xml
    zip.start_file("[Content_Types].xml", options)?;
    let mut content_types = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
"#,
    );
    for (i, _) in sheets.iter().enumerate() {
        let _ = write!(
            content_types,
            r#"<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
            i + 1
        );
    }
    content_types.push_str("</Types>");
    zip.write_all(content_types.as_bytes())?;

    // 2. _rels/.rels
    zip.start_file("_rels/.rels", options)?;
    zip.write_all(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#.as_bytes())?;

    // 3. xl/workbook.xml
    zip.start_file("xl/workbook.xml", options)?;
    let mut workbook_xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
"#,
    );
    for (i, name) in sheets.iter().enumerate() {
        let _ = write!(
            workbook_xml,
            r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
            name,
            i + 1,
            i + 1
        );
    }
    workbook_xml.push_str("</sheets>");

    if !ranges.is_empty() {
        workbook_xml.push_str("<definedNames>");
        for (name, content) in ranges {
            let _ = write!(
                workbook_xml,
                r#"<definedName name="{name}">{content}</definedName>"#
            );
        }
        workbook_xml.push_str("</definedNames>");
    }

    workbook_xml.push_str("</workbook>");
    zip.write_all(workbook_xml.as_bytes())?;

    // 4. xl/_rels/workbook.xml.rels
    zip.start_file("xl/_rels/workbook.xml.rels", options)?;
    let mut rels_xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#,
    );
    for (i, _) in sheets.iter().enumerate() {
        let _ = write!(
            rels_xml,
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>"#,
            i + 1,
            i + 1
        );
    }
    rels_xml.push_str("</Relationships>");
    zip.write_all(rels_xml.as_bytes())?;

    // 5. sheets
    for (i, _name) in sheets.iter().enumerate() {
        zip.start_file(format!("xl/worksheets/sheet{}.xml", i + 1), options)?;
        zip.write_all(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#.as_bytes())?;
    }

    zip.finish()?;
    Ok(())
}

#[test]
fn test_remove_sheets() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let input_path = dir.path().join("input.xlsx");
    let output_path = dir.path().join("output.xlsx");

    create_mock_xlsx(&input_path, &["Sheet1", "Sheet2", "Sheet3"], &[])?;

    let mut remove_sheets = HashSet::new();
    remove_sheets.insert("Sheet2".to_string());

    let mods = WorkbookModifications {
        remove_sheets: Some(remove_sheets),
        ..Default::default()
    };

    modify_workbook(&input_path, &output_path, &mods)?;

    // Verify
    let file = File::open(&output_path)?;
    let mut zip = zip::ZipArchive::new(file)?;

    // Check if files exist/removed in zip
    assert!(zip.by_name("xl/worksheets/sheet1.xml").is_ok());
    assert!(
        zip.by_name("xl/worksheets/sheet2.xml").is_err(),
        "Sheet2 file should be removed"
    );
    assert!(zip.by_name("xl/worksheets/sheet3.xml").is_ok());

    Ok(())
}

#[test]
fn test_remove_named_ranges() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let input_path = dir.path().join("input_ranges.xlsx");
    let output_path = dir.path().join("output_ranges.xlsx");

    create_mock_xlsx(
        &input_path,
        &["Sheet1"],
        &[("Range1", "Sheet1!$A$1"), ("Range2", "Sheet1!$B$1")],
    )?;

    let mut remove_ranges = HashSet::new();
    remove_ranges.insert("Range1".to_string());

    let mods = WorkbookModifications {
        remove_named_ranges: Some(remove_ranges),
        ..Default::default()
    };

    modify_workbook(&input_path, &output_path, &mods)?;

    // Verify by inspecting xml content (reading zip back usually harder for content,
    // but we can check if zip is valid and potentially extract workbook.xml)
    let file = File::open(&output_path)?;
    let mut zip = zip::ZipArchive::new(file)?;
    let mut workbook_file = zip.by_name("xl/workbook.xml")?;
    let mut content = String::new();
    std::io::Read::read_to_string(&mut workbook_file, &mut content)?;

    assert!(
        !content.contains(r#"name="Range1""#),
        "Range1 should be removed"
    );
    assert!(content.contains(r#"name="Range2""#), "Range2 should remain");

    Ok(())
}

#[test]
fn test_combined_operations() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let input_path = dir.path().join("input_combined.xlsx");
    let output_path = dir.path().join("output_combined.xlsx");

    create_mock_xlsx(
        &input_path,
        &["Sheet1", "DeleteMe"],
        &[
            ("KeepRange", "Sheet1!$A$1"),
            ("DeleteRange", "DeleteMe!$A$1"),
        ],
    )?;

    let mut remove_sheets = HashSet::new();
    remove_sheets.insert("DeleteMe".to_string());

    let mut remove_ranges = HashSet::new();
    remove_ranges.insert("DeleteRange".to_string());

    let mods = WorkbookModifications {
        remove_sheets: Some(remove_sheets),
        remove_named_ranges: Some(remove_ranges),
        ..Default::default()
    };

    modify_workbook(&input_path, &output_path, &mods)?;

    let file = File::open(&output_path)?;
    let mut zip = zip::ZipArchive::new(file)?;

    // Check sheet file removal
    assert!(zip.by_name("xl/worksheets/sheet1.xml").is_ok());
    assert!(zip.by_name("xl/worksheets/sheet2.xml").is_err()); // DeleteMe was 2nd

    // Check workbook.xml
    let mut workbook_file = zip.by_name("xl/workbook.xml")?;
    let mut content = String::new();
    std::io::Read::read_to_string(&mut workbook_file, &mut content)?;

    assert!(!content.contains("DeleteMe"));
    assert!(!content.contains("DeleteRange"));
    assert!(content.contains("KeepRange"));

    Ok(())
}

const SHEET_HEADER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#;

/// Formula cells covering every result type, a shared group (C2:C4), an array
/// formula (E2:E3) and a formula without a cached value (F2).
const FORMULA_SHEET: &str = r#"<sheetData>
<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="s"><v>0</v></c></row>
<row r="2"><c r="A2"><v>2</v></c><c r="B2" s="1"><f>A2*2</f><v>4</v></c><c r="C2"><f t="shared" ref="C2:C4" si="0">A2+$A$1</f><v>3</v></c><c r="D2" t="str"><f>"a&amp;b "&amp;A2</f><v>a&amp;b 2</v></c><c r="E2"><f t="array" ref="E2:E3">A2:A3*10</f><v>20</v></c><c r="F2"><f>A2+1</f></c></row>
<row r="3"><c r="A3"><v>3</v></c><c r="B3" t="b"><f>A3&gt;2</f><v>1</v></c><c r="C3"><f t="shared" si="0"/><v>4</v></c><c r="D3" t="e"><f>1/0</f><v>#DIV/0!</v></c><c r="E3"><v>30</v></c></row>
<row r="4"><c r="A4"><v>4</v></c><c r="C4"><f t="shared" si="0"/><v>5</v></c></row>
</sheetData></worksheet>"#;

/// Build an XLSX with one `Data` sheet holding `sheet_data`, a shared string
/// table and a calculation chain.
fn create_formula_xlsx(path: &Path, sheet_data: &str) -> anyhow::Result<()> {
    let file = File::create(path)?;
    let mut zip = ZipWriter::new(file);
    let options = SimpleFileOptions::default();
    let files = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>"#.to_string(),
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.to_string(),
        ),
        (
            "xl/workbook.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_string(),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>"#.to_string(),
        ),
        (
            "xl/sharedStrings.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>Label</t></si></sst>"#.to_string(),
        ),
        (
            "xl/calcChain.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B2" i="1"/></calcChain>"#.to_string(),
        ),
        ("xl/worksheets/sheet1.xml", format!("{SHEET_HEADER}{sheet_data}")),
    ];
    for (name, content) in files {
        zip.start_file(name, options)?;
        zip.write_all(content.as_bytes())?;
    }
    zip.finish()?;
    Ok(())
}

fn read_part(path: &Path, part: &str) -> anyhow::Result<String> {
    let mut zip = zip::ZipArchive::new(File::open(path)?)?;
    let mut content = String::new();
    std::io::Read::read_to_string(&mut zip.by_name(part)?, &mut content)?;
    Ok(content)
}

/// Convert the `Data` sheet of a fresh formula workbook within `scopes`.
fn convert_formula_xlsx(dir: &Path, scopes: &[&str]) -> anyhow::Result<std::path::PathBuf> {
    let input_path = dir.join("formulas.xlsx");
    let output_path = dir.join("values.xlsx");
    create_formula_xlsx(&input_path, FORMULA_SHEET)?;
    let mods = WorkbookModifications {
        convert_to_values: scopes.iter().map(|s| s.parse()).collect::<Result<_, _>>()?,
        ..Default::default()
    };
    modify_workbook(&input_path, &output_path, &mods)?;
    Ok(output_path)
}

fn formula_at(workbook: &sheetrs::reader::Workbook, cell: (u32, u32)) -> Option<String> {
    workbook.sheets[0].cells[&cell]
        .formula
        .as_deref()
        .map(str::to_string)
}

#[test]
fn test_convert_sheet_to_values() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let output_path = convert_formula_xlsx(dir.path(), &["Data"])?;

    let sheet = read_part(&output_path, "xl/worksheets/sheet1.xml")?;
    assert!(sheet.contains(r#"<c r="B2" s="1"><v>4</v></c>"#), "{sheet}");
    assert!(sheet.contains(r#"<c r="C2"><v>3</v></c>"#), "{sheet}");
    assert!(sheet.contains(r#"<c r="C3"><v>4</v></c>"#), "{sheet}");
    assert!(
        sheet.contains(
            r#"<c r="D2" t="inlineStr"><is><t xml:space="preserve">a&amp;b 2</t></is></c>"#
        ),
        "{sheet}"
    );
    assert!(sheet.contains(r#"<c r="E2"><v>20</v></c>"#), "{sheet}");
    assert!(sheet.contains(r#"<c r="B3" t="b"><v>1</v></c>"#), "{sheet}");
    assert!(
        sheet.contains(r#"<c r="D3" t="e"><v>#DIV/0!</v></c>"#),
        "{sheet}"
    );
    // Without a cached value there is nothing to convert to.
    assert!(sheet.contains(r#"<c r="F2"><f>A2+1</f></c>"#), "{sheet}");

    // The calculation chain would list cells that no longer hold formulas.
    let mut zip = zip::ZipArchive::new(File::open(&output_path)?)?;
    assert!(zip.by_name("xl/calcChain.xml").is_err());
    assert!(!read_part(&output_path, "[Content_Types].xml")?.contains("calcChain"));
    let rels = read_part(&output_path, "xl/_rels/workbook.xml.rels")?;
    assert!(!rels.contains("calcChain"));
    assert!(rels.contains("sharedStrings.xml"));

    let workbook = read_workbook(&output_path)?;
    let cells = &workbook.sheets[0].cells;
    assert_eq!(cells[&(1, 3)].value, CellValue::Text("a&b 2".into()));
    assert_eq!(cells[&(1, 1)].value, CellValue::Number(4.0));
    assert_eq!(cells[&(2, 1)].value, CellValue::Boolean(true));
    assert_eq!(cells[&(0, 1)].value, CellValue::Text("Label".into()));
    let formulas: Vec<_> = cells.values().filter(|c| c.formula.is_some()).collect();
    assert_eq!(formulas.len(), 1, "only F2 keeps its formula");
    Ok(())
}

#[test]
fn test_convert_rows_unshares_rest_of_group() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let output_path = convert_formula_xlsx(dir.path(), &["Data!2"])?;

    let sheet = read_part(&output_path, "xl/worksheets/sheet1.xml")?;
    assert!(sheet.contains(r#"<c r="C2"><v>3</v></c>"#), "{sheet}");
    assert!(
        sheet.contains(r#"<c r="C3"><f>A3+$A$1</f><v>4</v></c>"#),
        "{sheet}"
    );
    assert!(
        sheet.contains(r#"<c r="C4"><f>A4+$A$1</f><v>5</v></c>"#),
        "{sheet}"
    );
    // The array formula spans rows 2 and 3, so row 2 alone keeps it.
    assert!(
        sheet.contains(r#"<f t="array" ref="E2:E3">A2:A3*10</f>"#),
        "{sheet}"
    );
    // Other rows keep their formulas.
    assert!(sheet.contains("<f>A3&gt;2</f>"), "{sheet}");

    let workbook = read_workbook(&output_path)?;
    assert_eq!(formula_at(&workbook, (1, 1)), None);
    assert_eq!(formula_at(&workbook, (2, 2)).as_deref(), Some("A3+$A$1"));
    assert_eq!(formula_at(&workbook, (3, 2)).as_deref(), Some("A4+$A$1"));
    Ok(())
}

#[test]
fn test_convert_columns_keeps_shared_group_when_first_cell_is_kept() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let output_path = convert_formula_xlsx(dir.path(), &["Data!3:4", "Data!B"])?;

    let sheet = read_part(&output_path, "xl/worksheets/sheet1.xml")?;
    assert!(sheet.contains(r#"<c r="B2" s="1"><v>4</v></c>"#), "{sheet}");
    assert!(
        sheet.contains(r#"<f t="shared" ref="C2:C4" si="0">A2+$A$1</f>"#),
        "{sheet}"
    );
    assert!(sheet.contains(r#"<c r="C3"><v>4</v></c>"#), "{sheet}");
    assert!(sheet.contains(r#"<c r="C4"><v>5</v></c>"#), "{sheet}");
    assert!(
        sheet.contains(r#"<c r="D3" t="e"><v>#DIV/0!</v></c>"#),
        "{sheet}"
    );
    assert!(sheet.contains(r#"<f>"a&amp;b "&amp;A2</f>"#), "{sheet}");
    // Calculation chain is dropped whenever cells are converted.
    let mut zip = zip::ZipArchive::new(File::open(&output_path)?)?;
    assert!(zip.by_name("xl/calcChain.xml").is_err());
    Ok(())
}

#[test]
fn test_convert_unknown_sheet_fails() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let err = convert_formula_xlsx(dir.path(), &["Missing"]).unwrap_err();
    assert!(err.to_string().contains("Missing"), "{err}");
    Ok(())
}

#[test]
fn test_convert_removed_sheet_is_skipped() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let input_path = dir.path().join("input.xlsx");
    let output_path = dir.path().join("output.xlsx");
    create_mock_xlsx(&input_path, &["Sheet1", "Sheet2"], &[])?;

    let mods = WorkbookModifications {
        remove_sheets: Some(HashSet::from(["Sheet2".to_string()])),
        convert_to_values: vec!["Sheet2".parse()?],
        ..Default::default()
    };
    modify_workbook(&input_path, &output_path, &mods)?;

    let mut zip = zip::ZipArchive::new(File::open(&output_path)?)?;
    assert!(zip.by_name("xl/worksheets/sheet2.xml").is_err());
    Ok(())
}

/// Converting every sheet of the shared test asset keeps each cached value and
/// leaves no formula behind.
#[test]
fn test_convert_minimal_asset_keeps_values() -> anyhow::Result<()> {
    let input_path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../tests/minimal_test.xlsx");
    let dir = tempfile::tempdir()?;
    let output_path = dir.path().join("values.xlsx");

    let original = read_workbook(&input_path)?;
    let mods = WorkbookModifications {
        convert_to_values: original
            .sheets
            .iter()
            .map(|sheet| sheet.name.parse())
            .collect::<Result<_, _>>()?,
        ..Default::default()
    };
    modify_workbook(&input_path, &output_path, &mods)?;
    let converted = read_workbook(&output_path)?;

    let mut formula_cells = 0;
    for (before, after) in original.sheets.iter().zip(&converted.sheets) {
        assert_eq!(before.name, after.name);
        for (pos, cell) in &before.cells {
            let new_cell = &after.cells[pos];
            assert_eq!(new_cell.value, cell.value, "{}!{pos:?}", before.name);
            if cell.formula.is_some() {
                formula_cells += 1;
                assert_eq!(new_cell.formula, None, "{}!{pos:?}", before.name);
            }
        }
    }
    assert!(formula_cells > 0, "the asset should contain formulas");
    Ok(())
}

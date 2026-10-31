use sheetrs_core::reader::{CellValue, read_workbook};
use sheetrs_core::writer::{ValueScope, ValueTarget, WorkbookModifications, modify_workbook};
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

/// Write a ZIP archive holding `files`, as `(name, content)` pairs.
fn write_zip(path: &Path, files: &[(&str, String)]) -> anyhow::Result<()> {
    let mut zip = ZipWriter::new(File::create(path)?);
    for (name, content) in files {
        zip.start_file(*name, SimpleFileOptions::default())?;
        zip.write_all(content.as_bytes())?;
    }
    zip.finish()?;
    Ok(())
}

/// Build an XLSX with one `Data` sheet holding `sheet_data`, a shared string
/// table and a calculation chain.
fn create_formula_xlsx(path: &Path, sheet_data: &str) -> anyhow::Result<()> {
    write_zip(path, &[
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
    ])
}

fn read_part(path: &Path, part: &str) -> anyhow::Result<String> {
    let mut zip = zip::ZipArchive::new(File::open(path)?)?;
    let mut content = String::new();
    std::io::Read::read_to_string(&mut zip.by_name(part)?, &mut content)?;
    Ok(content)
}

/// Convert the `Data` sheet of a fresh formula workbook within `scopes`.
fn convert_formula_xlsx(dir: &Path, scopes: &[&str]) -> anyhow::Result<std::path::PathBuf> {
    convert_sheet_xlsx(dir, FORMULA_SHEET, scopes)
}

/// Convert, within `scopes`, the `Data` sheet of a workbook holding
/// `sheet_data`.
fn convert_sheet_xlsx(
    dir: &Path,
    sheet_data: &str,
    scopes: &[&str],
) -> anyhow::Result<std::path::PathBuf> {
    let input_path = dir.join("formulas.xlsx");
    let output_path = dir.join("values.xlsx");
    create_formula_xlsx(&input_path, sheet_data)?;
    let mods = WorkbookModifications {
        convert_to_values: scopes.iter().map(|s| s.parse()).collect::<Result<_, _>>()?,
        ..Default::default()
    };
    modify_workbook(&input_path, &output_path, &mods)?;
    Ok(output_path)
}

fn formula_at(workbook: &sheetrs_core::reader::Workbook, cell: (u32, u32)) -> Option<String> {
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

#[test]
fn test_convert_data_table_only_when_covered() -> anyhow::Result<()> {
    let sheet_data = r#"<sheetData>
<row r="1"><c r="A1"><v>1</v></c></row>
<row r="2"><c r="B2"><f t="dataTable" ref="B2:B3" dt2D="0" dtr="0" r1="A1"/><v>10</v></c></row>
<row r="3"><c r="B3"><v>20</v></c></row>
</sheetData></worksheet>"#;
    let dir = tempfile::tempdir()?;

    let output_path = convert_sheet_xlsx(dir.path(), sheet_data, &["Data!2"])?;
    let sheet = read_part(&output_path, "xl/worksheets/sheet1.xml")?;
    assert!(sheet.contains(r#"<f t="dataTable""#), "{sheet}");

    let output_path = convert_sheet_xlsx(dir.path(), sheet_data, &["Data!B"])?;
    let sheet = read_part(&output_path, "xl/worksheets/sheet1.xml")?;
    assert!(sheet.contains(r#"<c r="B2"><v>10</v></c>"#), "{sheet}");
    Ok(())
}

/// A dynamic array formula whose `#VALUE!` stands for a rich value, which the
/// reader reads as no value, keeps its formula.
#[test]
fn test_convert_keeps_array_error_placeholder() -> anyhow::Result<()> {
    let sheet_data = r#"<sheetData>
<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="e" cm="1" vm="1"><f t="array" ref="B1">A1:A2</f><v>#VALUE!</v></c><c r="C1" t="e"><f>A1/0</f><v>#DIV/0!</v></c></row>
</sheetData></worksheet>"#;
    let dir = tempfile::tempdir()?;
    let input_path = dir.path().join("formulas.xlsx");
    let output_path = convert_sheet_xlsx(dir.path(), sheet_data, &["Data"])?;

    let original = read_workbook(&input_path)?;
    let converted = read_workbook(&output_path)?;
    for cell in [(0, 1), (0, 2)] {
        assert_eq!(
            converted.sheets[0].cells[&cell].value, original.sheets[0].cells[&cell].value,
            "{cell:?}"
        );
    }
    assert_eq!(formula_at(&converted, (0, 1)).as_deref(), Some("A1:A2"));
    assert_eq!(formula_at(&converted, (0, 2)), None);
    Ok(())
}

/// A workbook whose sheet order, `sheetId` and part names all differ, as in
/// files saved by Excel after reordering sheets, with sheet-local names.
fn create_reordered_xlsx(path: &Path) -> anyhow::Result<()> {
    let sheet = |value: u32| {
        format!(
            r#"{SHEET_HEADER}<sheetData><row r="1"><c r="A1"><v>{value}</v></c></row></sheetData></worksheet>"#
        )
    };
    write_zip(path, &[
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_string(),
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.to_string(),
        ),
        (
            "xl/workbook.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView firstSheet="1" activeTab="2"/></bookViews><sheets><sheet name="First" sheetId="3" r:id="rId7"/><sheet name="Second" sheetId="1" r:id="rId5"/><sheet name="Third" sheetId="2" r:id="rId6"/></sheets><definedNames><definedName name="Local" localSheetId="0">First!$A$1</definedName><definedName name="Local" localSheetId="2">Third!$A$1</definedName><definedName name="Global">Third!$A$1</definedName></definedNames></workbook>"#.to_string(),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet1.xml"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#.to_string(),
        ),
        ("xl/worksheets/sheet1.xml", sheet(3)),
        ("xl/worksheets/sheet2.xml", sheet(1)),
        ("xl/worksheets/sheet3.xml", sheet(2)),
    ])
}

#[test]
fn test_remove_sheet_resolves_its_part_by_relationship() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let input_path = dir.path().join("reordered.xlsx");
    let output_path = dir.path().join("removed.xlsx");
    create_reordered_xlsx(&input_path)?;

    let mods = WorkbookModifications {
        remove_sheets: Some(HashSet::from(["First".to_string()])),
        ..Default::default()
    };
    modify_workbook(&input_path, &output_path, &mods)?;

    let mut zip = zip::ZipArchive::new(File::open(&output_path)?)?;
    assert!(zip.by_name("xl/worksheets/sheet2.xml").is_err());
    let rels = read_part(&output_path, "xl/_rels/workbook.xml.rels")?;
    assert!(!rels.contains("rId7"), "{rels}");
    let content_types = read_part(&output_path, "[Content_Types].xml")?;
    assert!(!content_types.contains("sheet2.xml"), "{content_types}");

    let workbook = read_workbook(&output_path)?;
    let values: Vec<_> = workbook
        .sheets
        .iter()
        .map(|sheet| (sheet.name.as_str(), sheet.cells[&(0, 0)].value.clone()))
        .collect();
    assert_eq!(
        values,
        [
            ("Second", CellValue::Number(2.0)),
            ("Third", CellValue::Number(3.0))
        ]
    );

    // Sheet indices move past the removed sheet; its local name goes with it.
    let workbook_xml = read_part(&output_path, "xl/workbook.xml")?;
    assert!(
        workbook_xml.contains(r#"<workbookView firstSheet="0" activeTab="1"/>"#),
        "{workbook_xml}"
    );
    assert!(
        workbook_xml.contains(
            r#"<definedName name="Local" localSheetId="1">Third!$A$1</definedName><definedName name="Global">"#
        ),
        "{workbook_xml}"
    );
    assert!(!workbook_xml.contains("First"), "{workbook_xml}");
    Ok(())
}

#[test]
fn test_remove_every_sheet_fails() -> anyhow::Result<()> {
    for asset in ["minimal_test.xlsx", "minimal_test.ods"] {
        let input_path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../tests")
            .join(asset);
        let dir = tempfile::tempdir()?;
        let sheets = read_workbook(&input_path)?
            .sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect();
        let mods = WorkbookModifications {
            remove_sheets: Some(sheets),
            ..Default::default()
        };
        let err = modify_workbook(&input_path, &dir.path().join(asset), &mods).unwrap_err();
        assert!(
            format!("{err:#}").contains("every sheet"),
            "{asset}: {err:#}"
        );
    }
    Ok(())
}

/// Convert `scopes` of a shared test asset and check, cell by cell, that
/// every value is kept, formulas inside the scopes are gone and formulas
/// outside them are untouched. Returns how many formulas were converted.
fn check_asset_conversion(asset: &str, scopes: &[&str]) -> anyhow::Result<usize> {
    let input_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../tests")
        .join(asset);
    let dir = tempfile::tempdir()?;
    let output_path = dir.path().join(asset);

    let scopes: Vec<ValueScope> = scopes.iter().map(|s| s.parse()).collect::<Result<_, _>>()?;
    let mods = WorkbookModifications {
        convert_to_values: scopes.clone(),
        ..Default::default()
    };
    modify_workbook(&input_path, &output_path, &mods)?;
    let original = read_workbook(&input_path)?;
    let converted = read_workbook(&output_path)?;

    let mut converted_formulas = 0;
    assert_eq!(original.sheets.len(), converted.sheets.len());
    for (before, after) in original.sheets.iter().zip(&converted.sheets) {
        assert_eq!(before.name, after.name);
        for ((row, col), cell) in &before.cells {
            let new_cell = &after.cells[&(row, col)];
            let at = format!("{asset} {}!{row}:{col}", before.name);
            assert_eq!(new_cell.value, cell.value, "{at}");
            let in_scope = scopes
                .iter()
                .any(|scope| scope.sheet == before.name && scope.contains(row, col));
            if cell.formula.is_some() && in_scope && !cell.is_array {
                assert_eq!(new_cell.formula, None, "{at}");
                converted_formulas += 1;
            } else if !in_scope {
                assert_eq!(new_cell.formula, cell.formula, "{at}");
            }
        }
    }
    Ok(converted_formulas)
}

/// Converting every sheet of the shared test assets keeps each cached value
/// and leaves no formula behind, in XLSX and ODS alike.
#[test]
fn test_convert_minimal_assets_whole_sheets() -> anyhow::Result<()> {
    let sheets =
        read_workbook(Path::new(env!("CARGO_MANIFEST_DIR")).join("../tests/minimal_test.xlsx"))?
            .sheets
            .iter()
            .map(|sheet| {
                ValueScope {
                    sheet: sheet.name.clone(),
                    target: ValueTarget::Sheet,
                }
                .to_string()
            })
            .collect::<Vec<_>>();
    let sheets: Vec<&str> = sheets.iter().map(String::as_str).collect();

    let xlsx = check_asset_conversion("minimal_test.xlsx", &sheets)?;
    let ods = check_asset_conversion("minimal_test.ods", &sheets)?;
    assert!(xlsx > 0, "the assets should contain formulas");
    assert_eq!(xlsx, ods, "XLSX and ODS should convert the same formulas");
    Ok(())
}

/// Column and row scopes convert the same cells in XLSX and ODS.
#[test]
fn test_convert_minimal_assets_columns_and_rows() -> anyhow::Result<()> {
    let scopes = ["Indexing tests!J:K", "Indexing tests!12:15", "Summary!D"];
    let xlsx = check_asset_conversion("minimal_test.xlsx", &scopes)?;
    let ods = check_asset_conversion("minimal_test.ods", &scopes)?;
    assert!(xlsx > 0, "the scopes should hold formulas");
    assert_eq!(xlsx, ods, "XLSX and ODS should convert the same formulas");
    Ok(())
}

#[test]
fn test_convert_ods_keeps_mimetype_first_and_stored() -> anyhow::Result<()> {
    let input_path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../tests/minimal_test.ods");
    let dir = tempfile::tempdir()?;
    let output_path = dir.path().join("values.ods");
    let mods = WorkbookModifications {
        convert_to_values: vec!["Summary".parse()?],
        ..Default::default()
    };
    modify_workbook(&input_path, &output_path, &mods)?;

    let mut zip = zip::ZipArchive::new(File::open(&output_path)?)?;
    let mimetype = zip.by_index(0)?;
    assert_eq!(mimetype.name(), "mimetype");
    assert_eq!(mimetype.compression(), zip::CompressionMethod::Stored);
    Ok(())
}

/// Remove sheets and convert formulas of another sheet, in one plan, on a
/// shared test asset: the kept sheets keep their order, values, names and
/// out-of-scope formulas. Returns the names of the kept sheets.
fn check_asset_plan(asset: &str) -> anyhow::Result<Vec<String>> {
    let input_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../tests")
        .join(asset);
    let dir = tempfile::tempdir()?;
    let output_path = dir.path().join(asset);
    let removed = HashSet::from(["Summary".to_string(), "empty".to_string()]);
    let scope: ValueScope = "Indexing tests!J:K".parse()?;
    let mods = WorkbookModifications {
        remove_sheets: Some(removed.clone()),
        convert_to_values: vec![scope.clone()],
        ..Default::default()
    };
    modify_workbook(&input_path, &output_path, &mods)?;
    let original = read_workbook(&input_path)?;
    let modified = read_workbook(&output_path)?;

    let kept: Vec<_> = original
        .sheets
        .iter()
        .filter(|sheet| !removed.contains(&sheet.name))
        .collect();
    assert_eq!(kept.len(), modified.sheets.len(), "{asset}");
    for (before, after) in kept.iter().zip(&modified.sheets) {
        assert_eq!(before.name, after.name, "{asset}");
        for ((row, col), cell) in &before.cells {
            let at = format!("{asset} {}!{row}:{col}", before.name);
            let new_cell = &after.cells[&(row, col)];
            assert_eq!(new_cell.value, cell.value, "{at}");
            if !(scope.sheet == before.name && scope.contains(row, col)) {
                assert_eq!(new_cell.formula, cell.formula, "{at}");
            }
        }
    }
    assert_eq!(modified.defined_names, original.defined_names, "{asset}");
    Ok(modified
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect())
}

/// Removing sheets and converting formulas in one plan keeps the same sheets
/// in XLSX and ODS.
#[test]
fn test_minimal_assets_plan_parity() -> anyhow::Result<()> {
    let xlsx = check_asset_plan("minimal_test.xlsx")?;
    let ods = check_asset_plan("minimal_test.ods")?;
    assert_eq!(xlsx, ods);
    Ok(())
}

#[test]
fn test_remove_sheets_and_ranges_ods() -> anyhow::Result<()> {
    let input_path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../tests/minimal_test.ods");
    let dir = tempfile::tempdir()?;
    let output_path = dir.path().join("removed.ods");
    let original = read_workbook(&input_path)?;
    let range = original
        .defined_names
        .keys()
        .next()
        .expect("the asset should define names")
        .clone();

    let mods = WorkbookModifications {
        remove_sheets: Some(HashSet::from(["empty".to_string()])),
        remove_named_ranges: Some(HashSet::from([range.clone()])),
        ..Default::default()
    };
    modify_workbook(&input_path, &output_path, &mods)?;

    let modified = read_workbook(&output_path)?;
    assert_eq!(modified.sheets.len(), original.sheets.len() - 1);
    assert!(modified.sheets.iter().all(|sheet| sheet.name != "empty"));
    assert!(!modified.defined_names.contains_key(&range));
    assert_eq!(
        modified.defined_names.len(),
        original.defined_names.len() - 1
    );
    Ok(())
}

#[test]
fn test_output_on_input_fails_and_keeps_input() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    std::fs::copy(
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../tests/minimal_test.xlsx"),
        &path,
    )?;
    let original = std::fs::read(&path)?;
    let mods = WorkbookModifications {
        convert_to_values: vec!["Summary".parse()?],
        ..Default::default()
    };

    let same_file = dir.path().join(".").join("book.xlsx");
    let err = modify_workbook(&path, &same_file, &mods).unwrap_err();
    assert!(err.to_string().contains("is the input file"), "{err}");
    assert_eq!(std::fs::read(&path)?, original);
    Ok(())
}

use sheetrs::writer::{WorkbookModifications, modify_workbook};
use std::collections::HashSet;
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
        content_types.push_str(&format!(
            r#"<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
            i + 1
        ));
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
        workbook_xml.push_str(&format!(
            r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
            name,
            i + 1,
            i + 1
        ));
    }
    workbook_xml.push_str("</sheets>");

    if !ranges.is_empty() {
        workbook_xml.push_str("<definedNames>");
        for (name, content) in ranges {
            workbook_xml.push_str(&format!(
                r#"<definedName name="{}">{}</definedName>"#,
                name, content
            ));
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
        rels_xml.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>"#,
            i + 1, i + 1
        ));
    }
    rels_xml.push_str("</Relationships>");
    zip.write_all(rels_xml.as_bytes())?;

    // 5. sheets
    for (i, _name) in sheets.iter().enumerate() {
        zip.start_file(&format!("xl/worksheets/sheet{}.xml", i + 1), options)?;
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
        remove_named_ranges: None,
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
        remove_sheets: None,
        remove_named_ranges: Some(remove_ranges),
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

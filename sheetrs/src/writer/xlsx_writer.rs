// ! XLSX writer functionality for removing sheets and named ranges

use crate::reader::parser_utils::attr_value;
use anyhow::Result;
use quick_xml::events::Event;
use quick_xml::{Reader, Writer};
use std::collections::HashSet;
use std::io::{Cursor, Read, Seek, Write};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

/// Struct used to define modifications to be applied to a workbook
#[derive(Debug, Default)]
pub struct WorkbookModifications {
    pub remove_sheets: Option<HashSet<String>>,
    pub remove_named_ranges: Option<HashSet<String>>,
}

/// Modify an XLSX file by applying specified modifications
pub fn modify_workbook_xlsx<R: Read + Seek, W: Write + Seek>(
    mut input: R,
    output: W,
    modifications: &WorkbookModifications,
) -> Result<()> {
    let mut archive = ZipArchive::new(&mut input)?;

    // Create output ZIP
    let mut zip_writer = ZipWriter::new(output);

    // Read workbook.xml needed for IDs and cleanup
    let workbook_xml = read_file_from_zip(&mut archive, "xl/workbook.xml")?;
    let sheet_info = parse_sheet_info(&workbook_xml)?;

    // Identify sheet IDs to remove
    let sheet_ids_to_remove: Vec<usize> = if let Some(sheets) = &modifications.remove_sheets {
        sheet_info
            .iter()
            .filter(|(name, _)| sheets.contains(name))
            .map(|(_, id)| *id)
            .collect()
    } else {
        Vec::new()
    };

    // Iterate through all files in the archive
    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let name = file.name().to_string();

        // Check if we should skip this file (e.g. removed sheet xml)
        let should_skip = sheet_ids_to_remove.iter().any(|id| {
            name == format!("xl/worksheets/sheet{id}.xml")
                || name == format!("xl/worksheets/_rels/sheet{id}.xml.rels")
        });

        if should_skip {
            continue;
        }

        // Processing logic
        if name == "xl/workbook.xml" {
            let mut content = workbook_xml.clone(); // Use cached content

            // 1. Remove sheets
            if let Some(sheets) = &modifications.remove_sheets
                && !sheets.is_empty()
            {
                content = remove_sheets_from_workbook_xml(&content, sheets)?;
            }

            // 2. Remove named ranges
            if let Some(ranges) = &modifications.remove_named_ranges
                && !ranges.is_empty()
            {
                content = remove_named_ranges_from_workbook_xml(&content, ranges)?;
            }

            zip_writer.start_file(&name, SimpleFileOptions::default())?;
            zip_writer.write_all(content.as_bytes())?;
        } else if name == "[Content_Types].xml" && !sheet_ids_to_remove.is_empty() {
            let mut content = String::new();
            file.read_to_string(&mut content)?;
            let modified_content = remove_sheet_content_types(&content, &sheet_ids_to_remove)?;
            zip_writer.start_file(&name, SimpleFileOptions::default())?;
            zip_writer.write_all(modified_content.as_bytes())?;
        } else if name == "xl/_rels/workbook.xml.rels" && !sheet_ids_to_remove.is_empty() {
            let mut content = String::new();
            file.read_to_string(&mut content)?;
            let modified_content = remove_sheet_relationships(&content, &sheet_ids_to_remove)?;
            zip_writer.start_file(&name, SimpleFileOptions::default())?;
            zip_writer.write_all(modified_content.as_bytes())?;
        } else {
            // Copy file as is
            zip_writer.start_file(&name, SimpleFileOptions::default())?;
            let mut buffer = Vec::new();
            file.read_to_end(&mut buffer)?;
            zip_writer.write_all(&buffer)?;
        }
    }

    zip_writer.finish()?;
    Ok(())
}

// Helper functions

fn read_file_from_zip<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    filename: &str,
) -> Result<String> {
    let mut file = archive.by_name(filename)?;
    let mut content = String::new();
    file.read_to_string(&mut content)?;
    Ok(content)
}

fn parse_sheet_info(workbook_xml: &str) -> Result<Vec<(String, usize)>> {
    let mut reader = Reader::from_str(workbook_xml);
    let mut buf = Vec::new();
    let mut sheets = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.name().as_ref() == "sheet" => {
                let mut name = String::new();
                let mut sheet_id = 0;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        "name" => {
                            name = attr_value(&attr)?.into_owned();
                        }
                        "sheetId" => {
                            sheet_id = attr_value(&attr)?.into_owned().parse()?;
                        }
                        _ => {}
                    }
                }

                sheets.push((name, sheet_id));
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("Error parsing XML: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(sheets)
}

fn remove_sheets_from_workbook_xml(
    xml: &str,
    sheets_to_remove: &HashSet<String>,
) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();
    let mut skip_current_sheet = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == "sheet" => {
                let mut sheet_name = String::new();
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == "name" {
                        sheet_name = attr_value(&attr)?.into_owned();
                        break;
                    }
                }

                if sheets_to_remove.contains(&sheet_name) {
                    skip_current_sheet = true;
                } else {
                    writer.write_event(Event::Start(e))?;
                }
            }
            Ok(Event::Empty(e)) if e.name().as_ref() == "sheet" => {
                let mut sheet_name = String::new();
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == "name" {
                        sheet_name = attr_value(&attr)?.into_owned();
                        break;
                    }
                }

                if !sheets_to_remove.contains(&sheet_name) {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Ok(Event::End(e)) if e.name().as_ref() == "sheet" => {
                if skip_current_sheet {
                    skip_current_sheet = false;
                } else {
                    writer.write_event(Event::End(e))?;
                }
            }
            Ok(Event::Eof) => break,
            Ok(e) => {
                if !skip_current_sheet {
                    writer.write_event(e)?;
                }
            }
            Err(e) => return Err(anyhow::anyhow!("Error parsing XML: {e}")),
        }
        buf.clear();
    }

    let result = writer.into_inner().into_inner();
    Ok(String::from_utf8(result)?)
}

fn remove_sheet_content_types(xml: &str, sheet_ids: &[usize]) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) if e.name().as_ref() == "Override" => {
                let mut part_name = String::new();
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == "PartName" {
                        part_name = attr_value(&attr)?.into_owned();
                        break;
                    }
                }

                // Skip if this is a sheet we're removing
                let should_skip = sheet_ids
                    .iter()
                    .any(|id| part_name == format!("/xl/worksheets/sheet{id}.xml"));

                if !should_skip {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Ok(Event::Eof) => break,
            Ok(e) => writer.write_event(e)?,
            Err(e) => return Err(anyhow::anyhow!("Error parsing XML: {e}")),
        }
        buf.clear();
    }

    let result = writer.into_inner().into_inner();
    Ok(String::from_utf8(result)?)
}

fn remove_sheet_relationships(xml: &str, sheet_ids: &[usize]) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) if e.name().as_ref() == "Relationship" => {
                let mut target = String::new();
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == "Target" {
                        target = attr_value(&attr)?.into_owned();
                        break;
                    }
                }

                // Skip if this is a sheet relationship we're removing
                let should_skip = sheet_ids
                    .iter()
                    .any(|id| target == format!("worksheets/sheet{id}.xml"));

                if !should_skip {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Ok(Event::Eof) => break,
            Ok(e) => writer.write_event(e)?,
            Err(e) => return Err(anyhow::anyhow!("Error parsing XML: {e}")),
        }
        buf.clear();
    }

    let result = writer.into_inner().into_inner();
    Ok(String::from_utf8(result)?)
}

fn remove_named_ranges_from_workbook_xml(
    xml: &str,
    ranges_to_remove: &HashSet<String>,
) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();
    let mut skip_current_range = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == "definedName" => {
                let mut range_name = String::new();
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == "name" {
                        range_name = attr_value(&attr)?.into_owned();
                        break;
                    }
                }

                if ranges_to_remove.contains(&range_name) {
                    skip_current_range = true;
                } else {
                    writer.write_event(Event::Start(e))?;
                }
            }
            Ok(Event::Empty(e)) if e.name().as_ref() == "definedName" => {
                let mut range_name = String::new();
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == "name" {
                        range_name = attr_value(&attr)?.into_owned();
                        break;
                    }
                }

                if !ranges_to_remove.contains(&range_name) {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Ok(Event::End(e)) if e.name().as_ref() == "definedName" => {
                if skip_current_range {
                    skip_current_range = false;
                } else {
                    writer.write_event(Event::End(e))?;
                }
            }
            Ok(Event::Eof) => break,
            Ok(e) => {
                if !skip_current_range {
                    writer.write_event(e)?;
                }
            }
            Err(e) => return Err(anyhow::anyhow!("Error parsing XML: {e}")),
        }
        buf.clear();
    }

    let result = writer.into_inner().into_inner();
    Ok(String::from_utf8(result)?)
}

#[cfg(test)]
mod tests {
    use super::*;

    const WORKBOOK_XML: &str = r#"<workbook><sheets><sheet name="R&amp;D" sheetId="1"/><sheet name="Data" sheetId="2"/></sheets><definedNames><definedName name="A&amp;B">Data!$A$1</definedName><definedName name="Keep">Data!$B$1</definedName></definedNames></workbook>"#;

    #[test]
    fn test_parse_sheet_info_unescapes_names() {
        let sheets = parse_sheet_info(WORKBOOK_XML).unwrap();
        assert_eq!(
            sheets,
            vec![("R&D".to_string(), 1), ("Data".to_string(), 2)]
        );
    }

    #[test]
    fn test_remove_sheet_with_escaped_name() {
        let remove = HashSet::from(["R&D".to_string()]);
        let xml = remove_sheets_from_workbook_xml(WORKBOOK_XML, &remove).unwrap();
        assert!(!xml.contains("R&amp;D"));
        assert!(xml.contains(r#"<sheet name="Data" sheetId="2"/>"#));
    }

    #[test]
    fn test_remove_named_range_with_escaped_name() {
        let remove = HashSet::from(["A&B".to_string()]);
        let xml = remove_named_ranges_from_workbook_xml(WORKBOOK_XML, &remove).unwrap();
        assert!(!xml.contains("A&amp;B"));
        assert!(xml.contains(r#"<definedName name="Keep">Data!$B$1</definedName>"#));
    }
}

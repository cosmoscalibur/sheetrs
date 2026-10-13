//! XML parsing utilities for extracting metadata from XLSX files

use anyhow::{Context, Result};
use quick_xml::Reader;
use quick_xml::events::Event;
use regex::{Captures, Regex, Replacer};
use std::collections::HashMap;
use std::fmt::Write;
use std::io::BufReader;
use std::sync::{Arc, LazyLock};
use zip::ZipArchive;

use super::parser_utils::{
    attr_value, column_letters, column_letters_to_index, parse_cell_range, parse_cell_ref,
    push_general_ref, read_text_node,
};
use super::{Cell, CellValue, Sheet, WorkbookReader};
use crate::violation::ExcelError;

/// Resolve sheet name to its XML path in the XLSX archive
pub fn get_xlsx_sheet_path(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    sheet_name: &str,
) -> Result<String> {
    // 1. Get rId from xl/workbook.xml
    let mut rid = String::new();
    {
        let workbook_xml = archive
            .by_name("xl/workbook.xml")
            .context("Failed to find xl/workbook.xml")?;
        let mut reader = Reader::from_reader(BufReader::new(workbook_xml));
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "sheet" => {
                    let mut name = String::new();
                    let mut r_id = String::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            "name" => name = attr_value(&attr)?.into_owned(),
                            "r:id" => r_id = attr_value(&attr)?.into_owned(),
                            _ => {}
                        }
                    }
                    if name == sheet_name {
                        rid = r_id;
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
    }

    if rid.is_empty() {
        return Err(anyhow::anyhow!(
            "Sheet '{sheet_name}' not found in workbook.xml"
        ));
    }

    // 2. Resolve rId in xl/_rels/workbook.xml.rels
    let mut target = String::new();
    {
        let rels_xml = archive
            .by_name("xl/_rels/workbook.xml.rels")
            .context("Failed to find xl/_rels/workbook.xml.rels")?;
        let mut reader = Reader::from_reader(BufReader::new(rels_xml));
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "Relationship" => {
                    let mut id = String::new();
                    let mut t = String::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            "Id" => id = attr_value(&attr)?.into_owned(),
                            "Target" => t = attr_value(&attr)?.into_owned(),
                            _ => {}
                        }
                    }
                    if id == rid {
                        target = t;
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
    }

    if target.is_empty() {
        return Err(anyhow::anyhow!(
            "Relationship '{rid}' not found for sheet '{sheet_name}'"
        ));
    }

    // Target is usually "worksheets/sheet1.xml", relative to "xl/", but may
    // also be an absolute part name such as "/xl/worksheets/sheet1.xml"
    if let Some(part) = target.strip_prefix('/') {
        Ok(part.to_string())
    } else if target.starts_with("worksheets/") {
        Ok(format!("xl/{target}"))
    } else {
        Ok(target)
    }
}

/// Extract defined names (named ranges) from XLSX file
pub fn extract_defined_names_from_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<HashMap<String, String>> {
    let mut defined_names = HashMap::new();

    // Try to read workbook.xml
    let Ok(workbook_xml) = archive.by_name("xl/workbook.xml") else {
        return Ok(defined_names); // No workbook.xml, return empty
    };

    let buf_reader = BufReader::new(workbook_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut in_defined_names = false;
    let mut current_name = String::new();
    let mut current_ref = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                match e.name().as_ref() {
                    "definedNames" => in_defined_names = true,
                    "definedName" if in_defined_names => {
                        // Get the name attribute
                        for attr in e.attributes() {
                            if let Ok(attr) = attr
                                && attr.key.as_ref() == "name"
                            {
                                current_name = attr_value(&attr)?.into_owned();
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(e)) if in_defined_names && !current_name.is_empty() => {
                current_ref.push_str(&e);
            }
            Ok(Event::GeneralRef(e)) if in_defined_names && !current_name.is_empty() => {
                push_general_ref(&mut current_ref, &e)?;
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                "definedName" if !current_name.is_empty() => {
                    // Filter out internal Excel names
                    if !current_name.starts_with("_xlnm.")
                        && !current_name.contains("_FilterDatabase")
                    {
                        defined_names.insert(current_name.clone(), current_ref.trim().to_string());
                    }
                    current_name.clear();
                    current_ref.clear();
                }
                "definedNames" => in_defined_names = false,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML parsing error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(defined_names)
}

/// Extract Excel Tables from XLSX file as defined names
/// Scans xl/tables/*.xml
pub fn extract_tables_from_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<HashMap<String, String>> {
    let mut current_tables = HashMap::new();
    let mut table_files = Vec::new();

    // Find all table files
    for i in 0..archive.len() {
        if let Ok(file) = archive.by_index(i) {
            let name = file.name().to_string();
            if name.starts_with("xl/tables/") && name.ends_with(".xml") {
                table_files.push(name);
            }
        }
    }

    for table_file in table_files {
        let Ok(xml) = archive.by_name(&table_file) else {
            continue;
        };

        let mut reader = Reader::from_reader(BufReader::new(xml));
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.name().as_ref() == "table" => {
                    let mut name = String::new();
                    let mut ref_sqref = String::new();

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                                "name" | "displayName"
                                    // displayName is usually the safe name, name might be id.
                                    // Spec says: name is collection name, displayName is unique name for formulas.
                                    // DisplayName is prioritized, but if name is empty, use either.
                                    if (name.is_empty() || attr.key.as_ref() == "displayName") => {
                                        name = attr_value(&attr)?.into_owned();
                                    }
                                "ref" => {
                                    ref_sqref = attr_value(&attr)?.into_owned();
                                }
                                _ => {}
                            }
                    }

                    if !name.is_empty() && !ref_sqref.is_empty() {
                        // Tables are usually local to the sheet they are in, but the table definition
                        // DOES NOT contain the sheet name in 'ref' (it's just A1:B2).
                        // However, defined names MUST include sheet name to be useful globally.
                        //
                        // The exact sheet ownership is not easily known from table XML alone without relationships.
                        // But Excel tables HAVE a unique name across the workbook.
                        // A formula refers to it by name `Table1`, not `Sheet1!Table1`.
                        // So storing just the name is sufficient.
                        //
                        // PERF001 checks for unused named ranges by looking for the name in formulas.
                        // If users use `=SUM(Table1)`, `extract_formulas` will return strings containing `Table1`.
                        // Registering `Table1` as a defined name enables this check.
                        // The associated value (range) is primarily for information/reporting.

                        current_tables.insert(name, ref_sqref);
                    }
                    // Only the top level table element is processed
                    break;
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }

    Ok(current_tables)
}

/// XLSX (Office Open XML) reader implementation
///
/// Handles parsing of .xlsx workbook.xml, worksheets/*.xml and other parts
pub struct XlsxReader<'a, R: std::io::Read + std::io::Seek> {
    archive: &'a mut ZipArchive<R>,
    shared_strings: Vec<Arc<str>>,
    styles: Vec<Arc<str>>,
}

impl<'a, R: std::io::Read + std::io::Seek> XlsxReader<'a, R> {
    pub fn new(archive: &'a mut ZipArchive<R>) -> Result<Self> {
        let shared_strings = extract_shared_strings(archive).unwrap_or_default();
        let styles = parse_styles(archive).unwrap_or_default();
        Ok(Self {
            archive,
            shared_strings,
            styles,
        })
    }
}

impl<'a, R: std::io::Read + std::io::Seek> WorkbookReader for XlsxReader<'a, R> {
    fn read_sheets(&mut self) -> Result<Vec<Sheet>> {
        let mut sheets = Vec::new();
        let sheet_names = self.get_sheet_names()?;
        let hidden_sheets = self.read_hidden_sheets()?;

        for (idx, name) in sheet_names.into_iter().enumerate() {
            let path = get_xlsx_sheet_path(self.archive, &name)?;
            let sheet_index = idx as u16;
            let mut sheet = Sheet::new(name.clone(), sheet_index);
            sheet.sheet_path = Some(path.clone());
            sheet.xml_size_bytes = self.archive.by_name(&path)?.size();
            sheet.visible = !hidden_sheets.contains(&name);

            // Parse sheet data
            let (cells, hidden_cols, hidden_rows, merged_cells, cf_count, cf_ranges, dim_range) =
                self.parse_sheet_xml(&path)?;

            sheet.cells = cells;
            sheet.hidden_columns = hidden_cols;
            sheet.hidden_rows = hidden_rows;
            sheet.merged_cells = merged_cells;
            sheet.conditional_formatting_count = cf_count;
            sheet.conditional_formatting_ranges = cf_ranges;

            sheet.used_range = dim_range;

            sheet.cells.shrink_to_fit();
            sheets.push(sheet);
        }

        Ok(sheets)
    }

    fn read_defined_names(&mut self) -> Result<HashMap<String, String>> {
        let mut names = extract_defined_names_from_xlsx(self.archive)?;
        let tables = extract_tables_from_xlsx(self.archive)?;
        names.extend(tables);
        Ok(names)
    }

    fn read_hidden_sheets(&mut self) -> Result<Vec<String>> {
        extract_hidden_sheets_from_xlsx(self.archive)
    }

    fn has_macros(&mut self) -> Result<bool> {
        // Check for vbaProject.bin, macrosheets, or any .bin files in xl/
        if self.archive.by_name("xl/vbaProject.bin").is_ok() {
            return Ok(true);
        }

        // Check for macrosheets
        for i in 0..self.archive.len() {
            if let Ok(file) = self.archive.by_index(i)
                && file.name().starts_with("xl/macrosheets/")
            {
                return Ok(true);
            }
        }

        Ok(false)
    }

    fn read_external_links(&mut self) -> Result<Vec<String>> {
        extract_external_links_xlsx(self.archive)
    }

    fn read_external_workbooks(&mut self) -> Result<Vec<super::ExternalWorkbook>> {
        extract_external_workbooks_xlsx(self.archive)
    }

    fn read_modified_date(&mut self) -> Result<Option<chrono::DateTime<chrono::Utc>>> {
        extract_modified_date_xlsx(self.archive)
    }

    fn read_date1904(&mut self) -> Result<bool> {
        extract_date1904_xlsx(self.archive)
    }
}

/// Extract last modified date from docProps/core.xml
fn extract_modified_date_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Option<chrono::DateTime<chrono::Utc>>> {
    use chrono::{DateTime, Utc};

    let Ok(core_xml) = archive.by_name("docProps/core.xml") else {
        return Ok(None);
    };

    let mut reader = Reader::from_reader(BufReader::new(core_xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut modified_date: Option<DateTime<Utc>> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                // dcterms:modified or cp:modified
                let name = e.name();
                if name.as_ref().ends_with(":modified") || name.as_ref() == "modified" {
                    let date_str = read_text_node(&mut reader)?;
                    // Parse ISO 8601 / RFC 3339 format
                    if let Ok(parsed) = DateTime::parse_from_rfc3339(date_str.trim()) {
                        modified_date = Some(parsed.with_timezone(&Utc));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(modified_date)
}

/// Extract date1904 setting from xl/workbook.xml
fn extract_date1904_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<bool> {
    let Ok(workbook_xml) = archive.by_name("xl/workbook.xml") else {
        return Ok(false);
    };

    let mut reader = Reader::from_reader(BufReader::new(workbook_xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "workbookPr" => {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == "date1904" {
                        let value = attr_value(&attr)?;
                        return Ok(value == "1" || value.eq_ignore_ascii_case("true"));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(false)
}

pub fn extract_external_links_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Vec<String>> {
    let mut links = Vec::new();
    let mut external_rels = Vec::new();

    // 1. Find external link relationships in workbook.xml.rels
    if let Ok(rels_xml) = archive.by_name("xl/_rels/workbook.xml.rels") {
        let mut reader = Reader::from_reader(BufReader::new(rels_xml));
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "Relationship" => {
                    let mut target = String::new();
                    let mut r_type = String::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            "Target" => target = attr_value(&attr)?.into_owned(),
                            "Type" => r_type = attr_value(&attr)?.into_owned(),
                            _ => {}
                        }
                    }
                    if r_type.ends_with("/externalLink") {
                        external_rels.push(target);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
    }

    // 2. Resolve each external link to its target workbook
    for rel_path in external_rels {
        // rel_path is usually like "externalLinks/externalLink1.xml"
        // We need the .rels for this file
        let filename = std::path::Path::new(&rel_path)
            .file_name()
            .and_then(|s| s.to_str())
            .unwrap_or("");

        let rels_of_ext = format!("xl/externalLinks/_rels/{filename}.rels");

        if let Ok(ext_rels_xml) = archive.by_name(&rels_of_ext) {
            let mut reader = Reader::from_reader(BufReader::new(ext_rels_xml));
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "Relationship" => {
                        let mut target = String::new();
                        let mut r_type = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "Target" => target = attr_value(&attr)?.into_owned(),
                                "Type" => r_type = attr_value(&attr)?.into_owned(),
                                _ => {}
                            }
                        }
                        // Accept both externalLinkPath (for file references) and externalWorkbook
                        if r_type.ends_with("/externalLinkPath")
                            || r_type.ends_with("/externalWorkbook")
                        {
                            links.push(target);
                        }
                    }
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }
    }

    Ok(links)
}

/// Extract external workbooks with indices from XLSX file
/// Returns a vector of ExternalWorkbook where index is 0-based (maps to [N+1] in formulas)
pub fn extract_external_workbooks_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Vec<super::ExternalWorkbook>> {
    use super::ExternalWorkbook;
    let mut workbooks = Vec::new();
    let mut external_rels = Vec::new();

    // 1. Find external link relationships in workbook.xml.rels
    // These are ordered and numbered (externalLink1.xml, externalLink2.xml, etc.)
    if let Ok(rels_xml) = archive.by_name("xl/_rels/workbook.xml.rels") {
        let mut reader = Reader::from_reader(BufReader::new(rels_xml));
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "Relationship" => {
                    let mut target = String::new();
                    let mut r_type = String::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            "Target" => target = attr_value(&attr)?.into_owned(),
                            "Type" => r_type = attr_value(&attr)?.into_owned(),
                            _ => {}
                        }
                    }
                    if r_type.ends_with("/externalLink") {
                        external_rels.push(target);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
    }

    // 2. Resolve each external link to its target workbook
    // Extract the numeric index from the filename (e.g., externalLink1.xml -> 1)
    for rel_path in external_rels {
        let filename = std::path::Path::new(&rel_path)
            .file_name()
            .and_then(|s| s.to_str())
            .unwrap_or("");

        // Extract index from filename (e.g., "externalLink1.xml" -> 1)
        // In XLSX, formulas use [1], [2], etc., so we convert to 0-based: index = N - 1
        let xlsx_index = filename
            .trim_start_matches("externalLink")
            .trim_end_matches(".xml")
            .parse::<usize>()
            .unwrap_or(1); // Default to 1 if parsing fails
        let index = xlsx_index.saturating_sub(1); // Convert to 0-based

        let rels_of_ext = format!("xl/externalLinks/_rels/{filename}.rels");

        if let Ok(ext_rels_xml) = archive.by_name(&rels_of_ext) {
            let mut reader = Reader::from_reader(BufReader::new(ext_rels_xml));
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "Relationship" => {
                        let mut target = String::new();
                        let mut r_type = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "Target" => target = attr_value(&attr)?.into_owned(),
                                "Type" => r_type = attr_value(&attr)?.into_owned(),
                                _ => {}
                            }
                        }
                        // Accept both externalLinkPath (for file references) and externalWorkbook
                        if r_type.ends_with("/externalLinkPath")
                            || r_type.ends_with("/externalWorkbook")
                        {
                            workbooks.push(ExternalWorkbook {
                                index,
                                path: super::parser_utils::extract_basename(&target),
                            });
                        }
                    }
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }
    }

    // Sort by index to ensure correct order
    workbooks.sort_by_key(|w| w.index);

    Ok(workbooks)
}

/// Cell reference inside a shared formula, optionally sheet-qualified.
///
/// Capture groups:
/// - 1: sheet qualifier (including `!`)
/// - 2: `$` marking an absolute column, or empty
/// - 3: column letters
/// - 4: `$` marking an absolute row, or empty
/// - 5: row number
///
/// Groups are numbered rather than named because a lookup by name hashes the
/// name on every match, and a shared formula is translated for every cell that
/// shares it.
static SHARED_FORMULA_REF_RE: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(r"((?:'[^']+'|[A-Za-z0-9_\.\-]+)!)?(\$?)([A-Z]+)(\$?)([0-9]+)")
        .expect("shared formula reference regex must compile")
});

/// `SHARED_FORMULA_REF_RE` replacer that writes each reference shifted by the
/// shared-formula offset straight into the output buffer.
struct SharedRefShifter {
    row_shift: i32,
    col_shift: i32,
}

impl Replacer for SharedRefShifter {
    fn replace_append(&mut self, caps: &Captures<'_>, dst: &mut String) {
        let col_abs = !caps[2].is_empty();
        let row_abs = !caps[4].is_empty();

        let col = column_letters_to_index(&caps[3]).expect("regex captures column letters");
        let row = caps[5].parse::<u32>().unwrap_or(1) - 1;

        let new_row = if row_abs {
            row
        } else {
            row.saturating_add_signed(self.row_shift)
        };
        let new_col = if col_abs {
            col
        } else {
            col.saturating_add_signed(self.col_shift)
        };

        if let Some(sheet) = caps.get(1) {
            dst.push_str(sheet.as_str());
        }
        if col_abs {
            dst.push('$');
        }
        let _ = write!(dst, "{}", column_letters(new_col));
        if row_abs {
            dst.push('$');
        }
        let _ = write!(dst, "{}", new_row + 1);
    }
}

pub(crate) fn translate_shared_formula(formula: &str, row_shift: i32, col_shift: i32) -> String {
    SHARED_FORMULA_REF_RE
        .replace_all(
            formula,
            SharedRefShifter {
                row_shift,
                col_shift,
            },
        )
        .into_owned()
}

// Type alias to avoid clippy::type_complexity warning
type ParsedSheetData = (
    HashMap<(u32, u32), Cell>,
    Vec<(u32, u32)>,
    Vec<(u32, u32)>,
    Vec<(u32, u32, u32, u32)>,
    usize,
    Vec<String>,
    Option<(u32, u32)>,
);

// Type alias to avoid clippy::type_complexity warning
type SharedFormulaMap = HashMap<u32, Vec<(String, u32, u32, Option<(u32, u32, u32, u32)>)>>;

impl<'a, R: std::io::Read + std::io::Seek> XlsxReader<'a, R> {
    fn get_sheet_names(&mut self) -> Result<Vec<String>> {
        let mut names = Vec::new();
        let workbook_xml = self.archive.by_name("xl/workbook.xml")?;
        let mut reader = Reader::from_reader(BufReader::new(workbook_xml));
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "sheet" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == "name" {
                            names.push(attr_value(&attr)?.into_owned());
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        Ok(names)
    }

    fn parse_sheet_xml(&mut self, path: &str) -> Result<ParsedSheetData> {
        let mut cells = HashMap::new();
        let mut hidden_columns = Vec::new();
        let mut hidden_rows = Vec::new();
        let mut merged_cells = Vec::new();
        let mut cf_count = 0;
        let mut cf_ranges = Vec::new();
        let mut dim_range = None;
        let mut shared_formulas: SharedFormulaMap = HashMap::new();

        let sheet_xml = self.archive.by_name(path)?;
        let mut reader = Reader::from_reader(BufReader::new(sheet_xml));
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut current_row = 0u32;
        let mut current_col = 0u32;

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => match e.name().as_ref() {
                    "dimension" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "ref" {
                                let ref_str = attr_value(&attr)?;
                                if let Some((_, _, end_row, end_col)) = parse_cell_range(&ref_str) {
                                    // defined range is inclusive, so count is index + 1
                                    dim_range = Some((end_row + 1, end_col + 1));
                                } else if let Some((end_row, end_col)) = parse_cell_ref(&ref_str) {
                                    // Single cell ref used as dimension? Rare but possible.
                                    dim_range = Some((end_row + 1, end_col + 1));
                                }
                            }
                        }
                    }
                    "col" => {
                        let mut min = 0u32;
                        let mut max = 0u32;
                        let mut hidden = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "min" => {
                                    min = attr_value(&attr)?.parse::<u32>()?.saturating_sub(1);
                                }
                                "max" => {
                                    max = attr_value(&attr)?.parse::<u32>()?.saturating_sub(1);
                                }
                                "hidden" => {
                                    hidden = attr.value == "1" || attr.value == "true";
                                }
                                _ => {}
                            }
                        }
                        if hidden && min <= max {
                            hidden_columns.push((min, max));
                        }
                    }
                    "row" => {
                        let mut hidden = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "r" => {
                                    current_row =
                                        attr_value(&attr)?.parse::<u32>()?.saturating_sub(1);
                                }
                                "hidden" => {
                                    hidden = attr.value == "1" || attr.value == "true";
                                }
                                _ => {}
                            }
                        }
                        if hidden {
                            hidden_rows.push((current_row, current_row));
                        }
                        current_col = 0;
                    }
                    "c" => {
                        let mut r_attr = String::new();
                        let mut s_attr = None;
                        let mut t_attr = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "r" => r_attr = attr_value(&attr)?.into_owned(),
                                "s" => s_attr = Some(attr_value(&attr)?.parse::<usize>()?),
                                "t" => t_attr = attr_value(&attr)?.into_owned(),
                                _ => {}
                            }
                        }

                        let (row, col) = if !r_attr.is_empty() {
                            let (r, c) =
                                parse_cell_ref(&r_attr).unwrap_or((current_row, current_col));
                            current_col = c + 1;
                            (r, c)
                        } else {
                            let c = current_col;
                            current_col += 1;
                            (current_row, c)
                        };

                        let num_fmt = s_attr.and_then(|idx| self.styles.get(idx).cloned());

                        let (value, mut formula, shared_si, shared_ref, is_array_formula) =
                            parse_cell_contents(
                                &mut reader,
                                &t_attr,
                                &self.shared_strings,
                                &self.styles,
                                num_fmt.as_deref(),
                            )?;

                        if let Some(si) = shared_si {
                            if let Some(f) = formula.as_ref() {
                                let range = shared_ref.and_then(|r| parse_cell_range(&r));
                                shared_formulas.entry(si).or_default().push((
                                    f.clone(),
                                    row,
                                    col,
                                    range,
                                ));
                            } else if let Some(defs) = shared_formulas.get(&si) {
                                // Find matching definition by range
                                let mut best_def = None;
                                for def in defs {
                                    if let Some((min_r, min_c, max_r, max_c)) = def.3
                                        && row >= min_r
                                        && row <= max_r
                                        && col >= min_c
                                        && col <= max_c
                                    {
                                        best_def = Some(def);
                                        break;
                                    }
                                }

                                // Fallback to last definition if no range match (or no range provided)
                                // This is risky but standard XML parsing often relies on latest def.
                                // However, strictly speaking, correct one should match range.
                                if best_def.is_none() {
                                    best_def = defs.last();
                                }

                                if let Some((base_formula, base_row, base_col, _)) = best_def
                                    && let Some(row_shift) = row.checked_signed_diff(*base_row)
                                    && let Some(col_shift) = col.checked_signed_diff(*base_col)
                                {
                                    formula = Some(translate_shared_formula(
                                        base_formula,
                                        row_shift,
                                        col_shift,
                                    ));
                                }
                            }
                        }

                        let mut cell = Cell {
                            row,
                            col,
                            value: value.clone(),
                            formula: None,
                            is_array: is_array_formula,
                            num_fmt,
                        };
                        if let Some(mut f) = formula {
                            if f.starts_with('=') {
                                f = f[1..].to_string();
                            }
                            cell.formula = Some(f.into());
                        }
                        cells.insert((row, col), cell);
                    }
                    "mergeCell" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "ref" {
                                let ref_str = attr_value(&attr)?;
                                if let Some(range) = parse_cell_range(&ref_str) {
                                    merged_cells.push(range);
                                }
                            }
                        }
                    }
                    "cfRule" => {
                        cf_count += 1;
                    }
                    "conditionalFormatting" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "sqref" {
                                let sqref = attr_value(&attr)?;
                                cf_ranges.push(sqref.to_string());
                            }
                        }
                    }

                    _ => {}
                },
                Event::Empty(e) => match e.name().as_ref() {
                    "dimension" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "ref" {
                                let ref_str = attr_value(&attr)?;
                                if let Some((_, _, end_row, end_col)) = parse_cell_range(&ref_str) {
                                    dim_range = Some((end_row + 1, end_col + 1));
                                } else if let Some((end_row, end_col)) = parse_cell_ref(&ref_str) {
                                    dim_range = Some((end_row + 1, end_col + 1));
                                }
                            }
                        }
                    }
                    "col" => {
                        let mut min = 0u32;
                        let mut max = 0u32;
                        let mut hidden = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "min" => {
                                    min = attr_value(&attr)?.parse::<u32>()?.saturating_sub(1);
                                }
                                "max" => {
                                    max = attr_value(&attr)?.parse::<u32>()?.saturating_sub(1);
                                }
                                "hidden" => {
                                    hidden = attr.value == "1" || attr.value == "true";
                                }
                                _ => {}
                            }
                        }
                        if hidden && min <= max {
                            hidden_columns.push((min, max));
                        }
                    }
                    "c" => {
                        let mut r_attr = String::new();
                        let mut s_attr = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "r" => r_attr = attr_value(&attr)?.into_owned(),
                                "s" => s_attr = Some(attr_value(&attr)?.parse::<usize>()?),
                                _ => {}
                            }
                        }
                        let (row, col) = if !r_attr.is_empty() {
                            let (r, c) =
                                parse_cell_ref(&r_attr).unwrap_or((current_row, current_col));
                            current_col = c + 1;
                            (r, c)
                        } else {
                            let c = current_col;
                            current_col += 1;
                            (current_row, c)
                        };
                        let num_fmt = s_attr.and_then(|idx| self.styles.get(idx).cloned());
                        cells.insert(
                            (row, col),
                            Cell {
                                row,
                                col,
                                value: CellValue::Empty,
                                formula: None,
                                is_array: false,
                                num_fmt,
                            },
                        );
                    }
                    "row" => {
                        // Empty row tag
                        let mut hidden = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "r" => {
                                    current_row =
                                        attr_value(&attr)?.parse::<u32>()?.saturating_sub(1);
                                }
                                "hidden" => {
                                    hidden = attr.value == "1" || attr.value == "true";
                                }
                                _ => {}
                            }
                        }
                        if hidden {
                            hidden_rows.push((current_row, current_row));
                        }
                        current_col = 0;
                    }
                    "mergeCell" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "ref" {
                                let ref_str = attr_value(&attr)?;
                                if let Some(range) = parse_cell_range(&ref_str) {
                                    merged_cells.push(range);
                                }
                            }
                        }
                    }
                    "cfRule" => {
                        cf_count += 1;
                    }
                    "conditionalFormatting" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "sqref" {
                                let sqref = attr_value(&attr)?;
                                // Strict normalization reverted to preserve original format (e.g. A1 instead of A1:A1):
                                cf_ranges.push(sqref.into_owned());
                            }
                        }
                    }
                    _ => {}
                },
                Event::End(e) if e.name().as_ref() == "worksheet" => {
                    break;
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        // Include hidden rows/columns in used_range for format parity
        // Both ODS and XLSX should report ALL empty rows/columns (visible or hidden) if they have properties
        if let Some((mut rows, mut cols)) = dim_range {
            if let Some(max_hidden_row) = hidden_rows.iter().map(|&(_, end)| end).max() {
                rows = rows.max(max_hidden_row + 1);
            }
            if let Some(max_hidden_col) = hidden_columns.iter().map(|&(_, end)| end).max() {
                cols = cols.max(max_hidden_col + 1);
            }
            dim_range = Some((rows, cols));
        } else if !hidden_rows.is_empty() || !hidden_columns.is_empty() {
            let mut rows = 1u32;
            let mut cols = 1u32;
            if let Some(max_hidden_row) = hidden_rows.iter().map(|&(_, end)| end).max() {
                rows = rows.max(max_hidden_row + 1);
            }
            if let Some(max_hidden_col) = hidden_columns.iter().map(|&(_, end)| end).max() {
                cols = cols.max(max_hidden_col + 1);
            }
            dim_range = Some((rows, cols));
        }

        Ok((
            cells,
            hidden_columns,
            hidden_rows,
            merged_cells,
            cf_count,
            cf_ranges,
            dim_range,
        ))
    }
}

// Type alias to avoid clippy::type_complexity warning
type ParsedCellData = (CellValue, Option<String>, Option<u32>, Option<String>, bool);

fn parse_cell_contents<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    t_attr: &str,
    shared_strings: &[Arc<str>],
    _styles: &[Arc<str>],
    num_fmt: Option<&str>,
) -> Result<ParsedCellData> {
    let mut value = CellValue::Empty;
    let mut formula = None;
    let mut shared_si = None;
    let mut shared_ref = None;
    let mut potential_error = None; // Store potential error value from t="e"
    let mut is_array_formula = false;
    let mut buf = Vec::new();

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => match e.name().as_ref() {
                "v" => {
                    let v_text = if let Event::Start(_) = event {
                        read_text_node(reader)?
                    } else {
                        String::new()
                    };
                    // Typed values ignore surrounding whitespace; text keeps it
                    let v_typed = v_text.trim();
                    value = match t_attr {
                        "s" => {
                            let idx = v_typed.parse::<usize>().unwrap_or(0);
                            CellValue::Text(
                                shared_strings
                                    .get(idx)
                                    .cloned()
                                    .unwrap_or_else(|| Arc::from("")),
                            )
                        }
                        "b" => CellValue::Boolean(v_typed == "1"),
                        "e" => {
                            // Store the error value but don't create error cell yet
                            // We need to check if there's a formula first
                            potential_error = Some(v_typed.to_string());
                            CellValue::Empty // Temporary, will be set later if needed
                        }
                        _ => {
                            // Check if this is a text-formatted number
                            // In XLSX, text format is indicated by num_fmt == "@"
                            if num_fmt == Some("@") {
                                // Store as text even if it looks like a number
                                CellValue::Text(Arc::from(v_text.as_str()))
                            } else if let Ok(n) = v_typed.parse::<f64>() {
                                CellValue::Number(n)
                            } else {
                                CellValue::Text(Arc::from(v_text.as_str()))
                            }
                        }
                    };
                }
                "f" => {
                    let mut si = None;
                    let mut is_shared = false;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            "si" => {
                                si = attr_value(&attr)?.parse::<u32>().ok();
                            }
                            "t" => match attr.value.as_ref() {
                                "shared" => is_shared = true,
                                "array" => is_array_formula = true,
                                _ => {}
                            },
                            "ref" => {
                                shared_ref = Some(attr_value(&attr)?.into_owned());
                            }
                            _ => {}
                        }
                    }

                    if let Event::Start(_) = event {
                        let f_text = read_text_node(reader)?;
                        if !f_text.is_empty() {
                            formula = Some(f_text);
                        }
                    }

                    if is_shared {
                        shared_si = si;
                    }
                }
                "is" => {
                    // Inline string can have multiple <t> tags
                    if let Event::Start(_) = event {
                        let mut is_text = String::new();
                        let mut is_buf = Vec::new();
                        loop {
                            match reader.read_event_into(&mut is_buf)? {
                                Event::Start(ref ee) if ee.name().as_ref() == "t" => {
                                    is_text.push_str(&read_text_node(reader)?);
                                }
                                Event::End(ref ee) if ee.name().as_ref() == "is" => break,
                                Event::Eof => break,
                                _ => {}
                            }
                            is_buf.clear();
                        }
                        value = CellValue::Text(Arc::from(is_text.as_str()));
                    }
                }
                _ => {}
            },
            Event::End(e) if e.name().as_ref() == "c" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    // Post-processing: handle potential_error
    // Excel uses t="e" for THREE different purposes:
    // 1. Real error cells (formula evaluated to error) - e.g., =car() -> #NAME?
    // 2. Array formulas (returns array, not error) - e.g., =Sheet!A1:B10 or =A1:A10
    // 3. Error constants, typed in or left by converting formulas to values
    //
    // The key distinction: array formulas contain range references (with ':')
    // Real errors are formulas without ranges that failed evaluation
    if let Some(err) = potential_error {
        // HEURISTIC: Legacy array formulas (CSE) often have t="e" and #VALUE! cached value
        // as a placeholder even if they aren't real errors. We ignore #VALUE! if it's an array.
        let array_placeholder = err == "#VALUE!"
            && formula
                .as_ref()
                .is_some_and(|f| is_array_formula || f.contains(':'));
        if !array_placeholder {
            value = CellValue::Error(ExcelError::from_cell_str(&err).unwrap_or(ExcelError::Value));
        }
    }

    Ok((value, formula, shared_si, shared_ref, is_array_formula))
}

pub fn extract_shared_strings(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Vec<Arc<str>>> {
    let mut strings = Vec::new();
    let Ok(ss_xml) = archive.by_name("xl/sharedStrings.xml") else {
        return Ok(strings);
    };

    let mut reader = Reader::from_reader(BufReader::new(ss_xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut current_string = String::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name().as_ref() == "t" => {
                current_string.push_str(&read_text_node(&mut reader)?);
            }
            Event::End(e) if e.name().as_ref() == "si" => {
                strings.push(Arc::from(current_string.as_str()));
                current_string.clear();
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(strings)
}

/// Count conditional formatting rules in a worksheet
pub fn count_conditional_formatting(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    sheet_name: &str,
) -> Result<usize> {
    let Ok(sheet_path) = get_xlsx_sheet_path(archive, sheet_name) else {
        return Ok(0);
    };

    let Ok(sheet_xml) = archive.by_name(&sheet_path) else {
        return Ok(0); // No sheet file, return 0
    };

    let buf_reader = BufReader::new(sheet_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut cf_count = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.name().as_ref() == "cfRule" => {
                cf_count += 1;
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(anyhow::anyhow!("XML parsing error: {e}"));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(cf_count)
}

/// Check if XLSX file contains VBA macros
pub fn has_vba_project_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<bool> {
    Ok(archive.by_name("xl/vbaProject.bin").is_ok())
}

/// Extract hidden sheets from XLSX file
pub fn extract_hidden_sheets_from_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Vec<String>> {
    let mut hidden_sheets = Vec::new();

    // Try to read workbook.xml
    let Ok(workbook_xml) = archive.by_name("xl/workbook.xml") else {
        return Ok(hidden_sheets); // No workbook.xml, return empty
    };

    let buf_reader = BufReader::new(workbook_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_sheets = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                match e.name().as_ref() {
                    "sheets" => in_sheets = true,
                    "sheet" if in_sheets => {
                        let mut name = String::new();
                        let mut state = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "name" => {
                                    name = attr_value(&attr)?.into_owned();
                                }
                                "state" => {
                                    state = attr_value(&attr)?.into_owned();
                                }
                                _ => {}
                            }
                        }

                        // If state is "hidden" or "veryHidden", add to list
                        if !name.is_empty() && (state == "hidden" || state == "veryHidden") {
                            hidden_sheets.push(name);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) if e.name().as_ref() == "sheets" => {
                in_sheets = false;
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML parsing error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(hidden_sheets)
}

// Type alias to avoid clippy::type_complexity warning
type HiddenColumnsRows = (Vec<(u32, u32)>, Vec<(u32, u32)>);

/// Extract hidden column and row ranges, (start, end) 0-based inclusive, from a worksheet
pub fn extract_hidden_columns_rows_from_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    sheet_name: &str,
) -> Result<HiddenColumnsRows> {
    let mut hidden_columns = Vec::new();
    let mut hidden_rows = Vec::new();

    let Ok(sheet_path) = get_xlsx_sheet_path(archive, sheet_name) else {
        return Ok((hidden_columns, hidden_rows));
    };

    let Ok(sheet_xml) = archive.by_name(&sheet_path) else {
        return Ok((hidden_columns, hidden_rows));
    };

    let buf_reader = BufReader::new(sheet_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                match e.name().as_ref() {
                    "col" => {
                        // Column definition with hidden attribute
                        let mut min_col = 0u32;
                        let mut max_col = 0u32;
                        let mut hidden = false;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "min" => {
                                    if let Ok(val) = attr_value(&attr)?.parse::<u32>() {
                                        min_col = val.saturating_sub(1); // Convert to 0-based
                                    }
                                }
                                "max" => {
                                    if let Ok(val) = attr_value(&attr)?.parse::<u32>() {
                                        max_col = val.saturating_sub(1); // Convert to 0-based
                                    }
                                }
                                "hidden" => {
                                    hidden = attr_value(&attr)? == "1"
                                        || attr_value(&attr)?.to_lowercase() == "true";
                                }
                                _ => {}
                            }
                        }

                        if hidden && min_col <= max_col {
                            hidden_columns.push((min_col, max_col));
                        }
                    }
                    "row" => {
                        // Row definition with hidden attribute
                        let mut row_num = 0u32;
                        let mut hidden = false;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "r" => {
                                    if let Ok(val) = attr_value(&attr)?.parse::<u32>() {
                                        row_num = val.saturating_sub(1); // Convert to 0-based
                                    }
                                }
                                "hidden" => {
                                    hidden = attr_value(&attr)? == "1"
                                        || attr_value(&attr)?.to_lowercase() == "true";
                                }
                                _ => {}
                            }
                        }

                        if hidden {
                            hidden_rows.push((row_num, row_num));
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML parsing error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok((hidden_columns, hidden_rows))
}

/// Extract merged cell ranges from a worksheet
pub fn extract_merged_cells_from_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    sheet_name: &str,
) -> Result<Vec<(u32, u32, u32, u32)>> {
    let mut merged_cells = Vec::new();

    let Ok(sheet_path) = get_xlsx_sheet_path(archive, sheet_name) else {
        return Ok(merged_cells);
    };

    let Ok(sheet_xml) = archive.by_name(&sheet_path) else {
        return Ok(merged_cells);
    };

    let buf_reader = BufReader::new(sheet_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.name().as_ref() == "mergeCell" => {
                for attr in e.attributes() {
                    if let Ok(attr) = attr
                        && attr.key.as_ref() == "ref"
                    {
                        let ref_str = attr_value(&attr)?;
                        if let Some((start_row, start_col, end_row, end_col)) =
                            parse_cell_range(&ref_str)
                        {
                            merged_cells.push((start_row, start_col, end_row, end_col));
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML parsing error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(merged_cells)
}

/// Built-in number formats by `numFmtId`, which `styles.xml` references without
/// declaring them.
///
/// See <https://github.com/dtjohnson/xlsx-populate/blob/master/lib/NumFmt.js>.
const BUILTIN_NUM_FMTS: &[(u32, &str)] = &[
    (0, "General"),
    (1, "0"),
    (2, "0.00"),
    (3, "#,##0"),
    (4, "#,##0.00"),
    (9, "0%"),
    (10, "0.00%"),
    (11, "0.00E+00"),
    (12, "# ?/?"),
    (13, "# ??/??"),
    (14, "mm-dd-yy"),
    (15, "d-mmm-yy"),
    (16, "d-mmm"),
    (17, "mmm-yy"),
    (18, "h:mm AM/PM"),
    (19, "h:mm:ss AM/PM"),
    (20, "h:mm"),
    (21, "h:mm:ss"),
    (22, "m/d/yy h:mm"),
    (37, "#,##0 ;(#,##0)"),
    (38, "#,##0 ;[Red](#,##0)"),
    (39, "#,##0.00;(#,##0.00)"),
    (40, "#,##0.00;[Red](#,##0.00)"),
    (45, "mm:ss"),
    (46, "[h]:mm:ss"),
    (47, "mmss.0"),
    (48, "##0.0E+0"),
    (49, "@"),
];

/// Parses `xl/styles.xml` into the number format code of each cell style (`xf`), by index.
pub fn parse_styles(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Vec<Arc<str>>> {
    let mut num_fmts: HashMap<u32, String> = BUILTIN_NUM_FMTS
        .iter()
        .map(|&(id, code)| (id, code.to_string()))
        .collect();

    let Ok(styles_xml) = archive.by_name("xl/styles.xml") else {
        return Ok(Vec::new());
    };

    let buf_reader = BufReader::new(styles_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut xfs = Vec::new();
    let mut in_cell_xfs = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                match e.name().as_ref() {
                    "numFmt" => {
                        let mut id = 0u32;
                        let mut code = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "numFmtId" => {
                                    if let Ok(val) = attr_value(&attr)?.parse::<u32>() {
                                        id = val;
                                    }
                                }
                                "formatCode" => {
                                    code = attr_value(&attr).unwrap_or_default().replace('\\', "");
                                }
                                _ => {}
                            }
                        }
                        if !code.is_empty() {
                            num_fmts.insert(id, code);
                        }
                    }
                    "cellXfs" => {
                        in_cell_xfs = true;
                    }
                    "xf" if in_cell_xfs => {
                        // Extract numFmtId
                        let mut num_fmt_id = 0u32;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr
                                && attr.key.as_ref() == "numFmtId"
                                && let Ok(val) = attr_value(&attr)?.parse::<u32>()
                            {
                                num_fmt_id = val;
                            }
                        }
                        // Look up format code and intern as Arc<str>
                        let format_code: Arc<str> = num_fmts
                            .get(&num_fmt_id)
                            .map(|s| Arc::from(s.as_str()))
                            .unwrap_or_else(|| Arc::from("General"));
                        xfs.push(format_code);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                match e.name().as_ref() {
                    "numFmt" => {
                        // Same as above, handle empty tag if it occurs (unlikely for numFmt with attrs)
                        let mut id = 0u32;
                        let mut code = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "numFmtId" => {
                                    if let Ok(val) = attr_value(&attr)?.parse::<u32>() {
                                        id = val;
                                    }
                                }
                                "formatCode" => {
                                    code = attr_value(&attr).unwrap_or_default().replace('\\', "");
                                }
                                _ => {}
                            }
                        }
                        if !code.is_empty() {
                            num_fmts.insert(id, code);
                        }
                    }
                    "xf" if in_cell_xfs => {
                        let mut num_fmt_id = 0u32;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr
                                && attr.key.as_ref() == "numFmtId"
                                && let Ok(val) = attr_value(&attr)?.parse::<u32>()
                            {
                                num_fmt_id = val;
                            }
                        }
                        let format_code: Arc<str> = num_fmts
                            .get(&num_fmt_id)
                            .map(|s| Arc::from(s.as_str()))
                            .unwrap_or_else(|| Arc::from("General"));
                        xfs.push(format_code);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) if e.name().as_ref() == "cellXfs" => {
                in_cell_xfs = false;
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(xfs)
}

/// Extract cell style indices from a worksheet
pub fn extract_cell_style_indices_from_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    sheet_name: &str,
) -> Result<HashMap<(u32, u32), usize>> {
    let mut cell_styles = HashMap::new();

    let Ok(sheet_path) = get_xlsx_sheet_path(archive, sheet_name) else {
        return Ok(cell_styles);
    };
    let Ok(sheet_xml) = archive.by_name(&sheet_path) else {
        return Ok(cell_styles);
    };

    let buf_reader = BufReader::new(sheet_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.name().as_ref() == "c" => {
                // Cell element `c`
                let mut row = 0u32;
                let mut col = 0u32;
                let mut style_index = 0usize;
                let mut has_style = false;

                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        "r" => {
                            let r_str = attr_value(&attr)?;
                            if let Some((r, c)) = parse_cell_ref(&r_str) {
                                row = r;
                                col = c;
                            }
                        }
                        "s" => {
                            if let Ok(val) = attr_value(&attr)?.parse::<usize>() {
                                style_index = val;
                                has_style = true;
                            }
                        }
                        _ => {}
                    }
                }

                if has_style {
                    cell_styles.insert((row, col), style_index);
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(cell_styles)
}

/// Extract formulas from an XLSX worksheet manually
pub fn extract_formulas_from_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    sheet_name: &str,
) -> Result<HashMap<(u32, u32), String>> {
    let mut formulas = HashMap::new();
    let mut shared_formulas = HashMap::new();

    let Ok(sheet_path) = get_xlsx_sheet_path(archive, sheet_name) else {
        return Ok(formulas);
    };

    let sheet_xml = archive.by_name(&sheet_path)?;
    let mut reader = Reader::from_reader(BufReader::new(sheet_xml));
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut current_cell_ref = String::new();
    let mut in_formula = false;
    let mut formula_text = String::new();

    loop {
        let event = reader.read_event_into(&mut buf)?;
        let is_start = matches!(event, Event::Start(_));
        match event {
            Event::Start(e) | Event::Empty(e) => match e.name().as_ref() {
                "c" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == "r" {
                            current_cell_ref = attr_value(&attr)?.into_owned();
                        }
                    }
                }
                "f" => {
                    // A self-closing <f/> has no text and no closing tag
                    in_formula = is_start;
                    formula_text.clear();
                    let mut si = None;
                    let mut t = None;

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            "si" => {
                                si = attr_value(&attr)?.parse::<u32>().ok();
                            }
                            "t" => {
                                t = Some(attr_value(&attr)?.into_owned());
                            }
                            _ => {}
                        }
                    }

                    if let Some(_s_idx) = si
                        && t.as_deref() == Some("shared")
                    {
                        // This might be the base or a consumer
                        // We wait for the text content to see if it's the base
                    }

                    // Store si for the current cell to association after getting Text
                    if let Some(s_idx) = si
                        && let Some((r, c)) = parse_cell_ref(&current_cell_ref)
                    {
                        // Temporarily store si to handle it in Text
                        // We use a prefix to distinguish from actual formulas
                        formulas.insert((r, c), format!("__SHARED__{s_idx}"));
                    }
                }
                _ => {}
            },
            Event::Text(e) if in_formula => formula_text.push_str(&e),
            Event::GeneralRef(e) if in_formula => push_general_ref(&mut formula_text, &e)?,
            Event::End(e) if e.name().as_ref() == "f" => {
                in_formula = false;
                let formula_text = std::mem::take(&mut formula_text);
                // An empty <f></f> is a shared-formula consumer: keep its marker
                if !formula_text.is_empty()
                    && let Some((r, c)) = parse_cell_ref(&current_cell_ref)
                {
                    // Check if this cell was marked as shared
                    if let Some(marker) = formulas.get(&(r, c))
                        && marker.starts_with("__SHARED__")
                    {
                        let si: u32 = marker["__SHARED__".len()..].parse().unwrap();
                        shared_formulas.insert(si, formula_text.clone());
                    }
                    formulas.insert((r, c), formula_text);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    // Second pass to resolve shared formulas
    for ((_r, _c), formula) in formulas.iter_mut() {
        if formula.starts_with("__SHARED__") {
            let si: u32 = formula["__SHARED__".len()..].parse().unwrap();
            if let Some(base_formula) = shared_formulas.get(&si) {
                *formula = base_formula.clone();
            } else {
                *formula = String::new(); // Fallback if si not found
            }
        }
    }

    Ok(formulas)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};

    #[test]
    fn test_translate_shared_formula_shifts_relative_refs() {
        assert_eq!(translate_shared_formula("A1+B2", 2, 1), "B3+C4");
        assert_eq!(translate_shared_formula("Z1", 0, 1), "AA1");
        assert_eq!(translate_shared_formula("Sheet2!C5", -1, -2), "Sheet2!A4");
        assert_eq!(
            translate_shared_formula("'My Sheet'!A1+1", 1, 0),
            "'My Sheet'!A2+1"
        );
    }

    #[test]
    fn test_translate_shared_formula_keeps_absolute_refs() {
        assert_eq!(
            translate_shared_formula("$A$1+A$1+$A1", 3, 3),
            "$A$1+D$1+$A4"
        );
    }

    #[test]
    fn test_translate_shared_formula_clamps_at_first_row_and_column() {
        assert_eq!(translate_shared_formula("B2", -5, -5), "A1");
    }

    fn parse_cell_value(cell_xml: &str, t_attr: &str, num_fmt: Option<&str>) -> CellValue {
        let mut reader = Reader::from_str(cell_xml);
        reader.read_event().unwrap();
        parse_cell_contents(&mut reader, t_attr, &[], &[], num_fmt)
            .unwrap()
            .0
    }

    #[test]
    fn test_parse_cell_contents_types_value_by_t_attribute() {
        assert_eq!(
            parse_cell_value("<c><v>42</v></c>", "", None),
            CellValue::Number(42.0)
        );
        assert_eq!(
            parse_cell_value("<c><v>1011</v></c>", "", Some("@")),
            CellValue::Text(Arc::from("1011"))
        );
        assert_eq!(
            parse_cell_value("<c><v>1</v></c>", "b", None),
            CellValue::Boolean(true)
        );
        assert_eq!(
            parse_cell_value("<c><v>0</v></c>", "b", None),
            CellValue::Boolean(false)
        );
    }

    #[test]
    fn test_parse_cell_contents_trims_typed_values() {
        assert_eq!(
            parse_cell_value("<c><v> 42 </v></c>", "", None),
            CellValue::Number(42.0)
        );
        assert_eq!(
            parse_cell_value("<c><v>\n 1\n</v></c>", "b", None),
            CellValue::Boolean(true)
        );
        assert_eq!(
            parse_cell_value("<c><v> a b </v></c>", "str", None),
            CellValue::Text(Arc::from(" a b "))
        );
    }

    #[test]
    fn test_parse_cell_contents_reads_error_constants() {
        assert_eq!(
            parse_cell_value("<c><v>#DIV/0!</v></c>", "e", None),
            CellValue::Error(ExcelError::DivZero)
        );
        // Legacy array formulas store #VALUE! as a placeholder result.
        assert_eq!(
            parse_cell_value(
                r#"<c><f t="array" ref="A1:A2">B1:B2</f><v>#VALUE!</v></c>"#,
                "e",
                None
            ),
            CellValue::Empty
        );
        assert_eq!(
            parse_cell_value("<c><f>1/0</f><v>#DIV/0!</v></c>", "e", None),
            CellValue::Error(ExcelError::DivZero)
        );
    }

    /// Build an XLSX archive holding one sheet, "Data", with `sheet_xml` as
    /// its worksheet part and `defined_names` inside `<definedNames>`.
    fn single_sheet_archive(sheet_xml: &str, defined_names: &str) -> ZipArchive<Cursor<Vec<u8>>> {
        use zip::write::SimpleFileOptions;

        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options =
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
            let parts = [
                (
                    "xl/workbook.xml",
                    format!(
                        r#"<workbook xmlns:r="r"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><definedNames>{defined_names}</definedNames></workbook>"#
                    ),
                ),
                (
                    "xl/_rels/workbook.xml.rels",
                    r#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>"#
                        .to_string(),
                ),
                ("xl/worksheets/sheet1.xml", sheet_xml.to_string()),
            ];
            for (name, content) in parts {
                zip.start_file(name, options).unwrap();
                zip.write_all(content.as_bytes()).unwrap();
            }
            zip.finish().unwrap();
        }
        ZipArchive::new(Cursor::new(buf)).unwrap()
    }

    #[test]
    fn test_extract_formulas_resolves_shared_consumers() {
        let mut archive = single_sheet_archive(
            r#"<worksheet><sheetData><row r="1">
                <c r="A1"><f t="shared" si="0" ref="A1:A3">B1&amp;&quot;x&quot;</f><v>1</v></c>
                <c r="A2"><f t="shared" si="0"></f><v>2</v></c>
                <c r="A3"><f t="shared" si="0"/><v>3</v></c>
                <c r="A4"><f>C4 &lt; 0</f><v>0</v></c>
            </row></sheetData></worksheet>"#,
            "",
        );

        let formulas = extract_formulas_from_xlsx(&mut archive, "Data").unwrap();

        assert_eq!(formulas.get(&(0, 0)).map(String::as_str), Some(r#"B1&"x""#));
        assert_eq!(formulas.get(&(1, 0)).map(String::as_str), Some(r#"B1&"x""#));
        assert_eq!(formulas.get(&(2, 0)).map(String::as_str), Some(r#"B1&"x""#));
        assert_eq!(formulas.get(&(3, 0)).map(String::as_str), Some("C4 < 0"));
    }

    #[test]
    fn test_extract_defined_names_trims_and_resolves_references() {
        let mut archive = single_sheet_archive(
            "<worksheet/>",
            "<definedName name=\"Range\">\n  &apos;My Data&apos;!$A$1:$B$2\n</definedName>",
        );

        let names = extract_defined_names_from_xlsx(&mut archive).unwrap();

        assert_eq!(
            names.get("Range").map(String::as_str),
            Some("'My Data'!$A$1:$B$2")
        );
    }

    #[test]
    fn test_extract_tables_from_xlsx() {
        use std::io::Cursor;
        use zip::write::SimpleFileOptions;

        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));

            let options =
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);

            // Add a table file
            zip.start_file("xl/tables/table1.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="MyTable" ref="A1:C3" tableType="xml" headerRowCount="1">
    <tableColumns count="3">
        <tableColumn id="1" name="Col1"/>
        <tableColumn id="2" name="Col2"/>
        <tableColumn id="3" name="Col3"/>
    </tableColumns>
</table>"#).unwrap();

            // Add another table file (without displayName)
            zip.start_file("xl/tables/table2.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="OtherTable" ref="D4:E5" tableType="xml" headerRowCount="1">
</table>"#).unwrap();

            zip.finish().unwrap();
        }

        let mut archive = ZipArchive::new(Cursor::new(buf)).unwrap();
        let tables = extract_tables_from_xlsx(&mut archive).unwrap();

        assert_eq!(tables.len(), 2);
        assert_eq!(tables.get("MyTable"), Some(&"A1:C3".to_string()));
        assert_eq!(tables.get("OtherTable"), Some(&"D4:E5".to_string()));
    }

    #[test]
    fn test_hidden_rows_and_columns_as_ranges_xlsx() {
        use std::io::Cursor;
        use zip::write::SimpleFileOptions;

        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options =
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);

            zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1"/>
    <cols>
        <col min="1" max="3" hidden="1"/>
        <col min="4" max="16384"/>
    </cols>
    <sheetData>
        <row r="1"><c r="A1"><v>1</v></c></row>
        <row r="2" hidden="1"/>
        <row hidden="1" r="3"></row>
        <row hidden="1" r="1048576"/>
    </sheetData>
</worksheet>"#,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let mut archive = ZipArchive::new(Cursor::new(buf)).unwrap();
        let mut reader = XlsxReader::new(&mut archive).unwrap();
        let (_, hidden_columns, hidden_rows, _, _, _, used_range) =
            reader.parse_sheet_xml("xl/worksheets/sheet1.xml").unwrap();

        assert_eq!(hidden_columns, vec![(0, 2)]);
        assert_eq!(hidden_rows, vec![(1, 1), (2, 2), (1_048_575, 1_048_575)]);
        assert_eq!(used_range, Some((1_048_576, 3)));
    }
}

#[test]
fn test_extract_external_workbooks_with_test_asset() {
    const TEST_XLSX: &[u8] = include_bytes!("../../../tests/minimal_test.xlsx");
    let cursor = std::io::Cursor::new(TEST_XLSX);
    let mut archive = ZipArchive::new(cursor).unwrap();

    let workbooks = extract_external_workbooks_xlsx(&mut archive).unwrap();

    // Verify external workbooks are extracted
    assert!(
        !workbooks.is_empty(),
        "Should detect external workbooks in test file"
    );

    // Verify indices are 0-based and sequential
    for (i, wb) in workbooks.iter().enumerate() {
        assert_eq!(wb.index, i, "Indices should be sequential 0-based");
    }

    // Verify paths are not empty
    for wb in &workbooks {
        assert!(
            !wb.path.is_empty(),
            "External workbook path should not be empty"
        );
    }
}

#[test]
fn test_sheet_collection_xlsx() {
    const TEST_XLSX: &[u8] = include_bytes!("../../../tests/minimal_test.xlsx");
    let cursor = std::io::Cursor::new(TEST_XLSX);
    let mut archive = ZipArchive::new(cursor).unwrap();
    let mut reader = XlsxReader::new(&mut archive).unwrap();

    let sheets = reader.read_sheets().unwrap();

    // Verify sheet count (should not include external sheets)
    assert!(!sheets.is_empty(), "Should have at least one sheet");

    // Verify no external sheet references in names
    for sheet in &sheets {
        assert!(
            !sheet.name.contains("file:///"),
            "Sheet collection should not contain external sheets: {}",
            sheet.name
        );
    }

    // Verify expected sheets are present
    let sheet_names: Vec<&str> = sheets.iter().map(|s| s.name.as_str()).collect();
    assert!(
        sheet_names.contains(&"Sheet7") || sheet_names.contains(&"Indexing tests"),
        "Should contain expected sheets"
    );
}

#[test]
fn test_sheet_visibility_xlsx() {
    const TEST_XLSX: &[u8] = include_bytes!("../../../tests/minimal_test.xlsx");
    let cursor = std::io::Cursor::new(TEST_XLSX);
    let mut archive = ZipArchive::new(cursor).unwrap();
    let mut reader = XlsxReader::new(&mut archive).unwrap();

    let sheets = reader.read_sheets().unwrap();

    // Count visible and hidden sheets
    let visible_count = sheets.iter().filter(|s| s.visible).count();
    let hidden_count = sheets.iter().filter(|s| !s.visible).count();

    assert!(visible_count > 0, "Should have at least one visible sheet");
    assert!(hidden_count > 0, "Test file should have hidden sheets");

    // Verify hidden sheets have visible=false
    for sheet in &sheets {
        if sheet.name.contains("hidden") || sheet.name.contains("empty_hidden") {
            assert!(!sheet.visible, "Sheet '{}' should be hidden", sheet.name);
        }
    }
}

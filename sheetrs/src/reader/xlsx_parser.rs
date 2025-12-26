//! XML parsing utilities for extracting metadata from XLSX files

use anyhow::{Context, Result};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashMap;
use std::io::BufReader;
use zip::ZipArchive;

use super::{Cell, CellValue, Sheet, WorkbookReader};

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
                Event::Start(e) | Event::Empty(e) => {
                    if e.name().as_ref() == b"sheet" {
                        let mut name = String::new();
                        let mut r_id = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"name" => name = attr.unescape_value()?.to_string(),
                                b"r:id" => r_id = attr.unescape_value()?.to_string(),
                                _ => {}
                            }
                        }
                        if name == sheet_name {
                            rid = r_id;
                            break;
                        }
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
            "Sheet '{}' not found in workbook.xml",
            sheet_name
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
                Event::Start(e) | Event::Empty(e) => {
                    if e.name().as_ref() == b"Relationship" {
                        let mut id = String::new();
                        let mut t = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"Id" => id = attr.unescape_value()?.to_string(),
                                b"Target" => t = attr.unescape_value()?.to_string(),
                                _ => {}
                            }
                        }
                        if id == rid {
                            target = t;
                            break;
                        }
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
            "Relationship '{}' not found for sheet '{}'",
            rid,
            sheet_name
        ));
    }

    // Target is usually "worksheets/sheet1.xml", we need to prepend "xl/" if it's relative
    if target.starts_with("worksheets/") {
        Ok(format!("xl/{}", target))
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
    let workbook_xml = match archive.by_name("xl/workbook.xml") {
        Ok(file) => file,
        Err(_) => return Ok(defined_names), // No workbook.xml, return empty
    };

    let buf_reader = BufReader::new(workbook_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_defined_names = false;
    let mut current_name = String::new();
    let mut current_ref = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                match e.name().as_ref() {
                    b"definedNames" => in_defined_names = true,
                    b"definedName" if in_defined_names => {
                        // Get the name attribute
                        for attr in e.attributes() {
                            if let Ok(attr) = attr
                                && attr.key.as_ref() == b"name"
                            {
                                current_name = attr.unescape_value()?.to_string();
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(e)) if in_defined_names && !current_name.is_empty() => {
                current_ref = e.unescape().unwrap_or_default().to_string();
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"definedName" if !current_name.is_empty() => {
                    // Filter out internal Excel names
                    if !current_name.starts_with("_xlnm.")
                        && !current_name.contains("_FilterDatabase")
                    {
                        defined_names.insert(current_name.clone(), current_ref.clone());
                    }
                    current_name.clear();
                    current_ref.clear();
                }
                b"definedNames" => in_defined_names = false,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML parsing error: {}", e)),
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
        let xml = match archive.by_name(&table_file) {
            Ok(file) => file,
            Err(_) => continue,
        };

        let mut reader = Reader::from_reader(BufReader::new(xml));
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"table" {
                        let mut name = String::new();
                        let mut ref_sqref = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"name" | b"displayName" => {
                                    // displayName is usually the safe name, name might be id.
                                    // Spec says: name is collection name, displayName is unique name for formulas.
                                    // DisplayName is prioritized.
                                    if name.is_empty() {
                                        name = attr.unescape_value()?.to_string();
                                    } else if attr.key.as_ref() == b"displayName" {
                                        name = attr.unescape_value()?.to_string();
                                    }
                                }
                                b"ref" => {
                                    ref_sqref = attr.unescape_value()?.to_string();
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

pub struct XlsxReader<'a, R: std::io::Read + std::io::Seek> {
    archive: &'a mut ZipArchive<R>,
    shared_strings: Vec<String>,
    styles: Vec<String>,
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

        for name in sheet_names {
            let path = get_xlsx_sheet_path(self.archive, &name)?;
            let mut sheet = Sheet::new(name.clone());
            sheet.sheet_path = Some(path.clone());
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
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                    let mut target = String::new();
                    let mut r_type = String::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"Target" => target = attr.unescape_value()?.to_string(),
                            b"Type" => r_type = attr.unescape_value()?.to_string(),
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

        let rels_of_ext = format!("xl/externalLinks/_rels/{}.rels", filename);

        if let Ok(ext_rels_xml) = archive.by_name(&rels_of_ext) {
            let mut reader = Reader::from_reader(BufReader::new(ext_rels_xml));
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                        let mut target = String::new();
                        let mut r_type = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"Target" => target = attr.unescape_value()?.to_string(),
                                b"Type" => r_type = attr.unescape_value()?.to_string(),
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
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                    let mut target = String::new();
                    let mut r_type = String::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"Target" => target = attr.unescape_value()?.to_string(),
                            b"Type" => r_type = attr.unescape_value()?.to_string(),
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

        let rels_of_ext = format!("xl/externalLinks/_rels/{}.rels", filename);

        if let Ok(ext_rels_xml) = archive.by_name(&rels_of_ext) {
            let mut reader = Reader::from_reader(BufReader::new(ext_rels_xml));
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                        let mut target = String::new();
                        let mut r_type = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"Target" => target = attr.unescape_value()?.to_string(),
                                b"Type" => r_type = attr.unescape_value()?.to_string(),
                                _ => {}
                            }
                        }
                        // Accept both externalLinkPath (for file references) and externalWorkbook
                        if r_type.ends_with("/externalLinkPath")
                            || r_type.ends_with("/externalWorkbook")
                        {
                            workbooks.push(ExternalWorkbook {
                                index,
                                path: target,
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

fn translate_shared_formula(formula: &str, row_shift: i32, col_shift: i32) -> String {
    thread_local! {
        static RE: regex::Regex = regex::Regex::new(r"(?P<sheet>(?:'[^']+'|[A-Za-z0-9_\.\-]+)!)?(?P<col_abs>\$?)(?P<col>[A-Z]+)(?P<row_abs>\$?)(?P<row>[0-9]+)").unwrap();
    }
    RE.with(|re| {
        re.replace_all(formula, |caps: &regex::Captures| {
            let sheet = caps.name("sheet").map(|m| m.as_str()).unwrap_or("");
            let col_abs = !caps.name("col_abs").unwrap().as_str().is_empty();
            let col_str = caps.name("col").unwrap().as_str();
            let row_abs = !caps.name("row_abs").unwrap().as_str().is_empty();
            let row_str = caps.name("row").unwrap().as_str();

            let mut col = 0u32;
            for ch in col_str.chars() {
                col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
            }
            col -= 1;

            let row = row_str.parse::<u32>().unwrap_or(1) - 1;

            let new_row = if row_abs {
                row
            } else {
                (row as i32 + row_shift).max(0) as u32
            };
            let new_col = if col_abs {
                col
            } else {
                (col as i32 + col_shift).max(0) as u32
            };

            let mut result = sheet.to_string();
            if col_abs {
                result.push('$');
            }

            // Col to letter logic
            let mut c = new_col + 1;
            let mut col_letter = String::new();
            while c > 0 {
                let m = (c - 1) % 26;
                col_letter.insert(0, (b'A' + m as u8) as char);
                c = (c - m) / 26;
            }
            result.push_str(&col_letter);

            if row_abs {
                result.push('$');
            }
            result.push_str(&(new_row + 1).to_string());

            result
        })
        .to_string()
    })
}

impl<'a, R: std::io::Read + std::io::Seek> XlsxReader<'a, R> {
    fn get_sheet_names(&mut self) -> Result<Vec<String>> {
        let mut names = Vec::new();
        let workbook_xml = self.archive.by_name("xl/workbook.xml")?;
        let mut reader = Reader::from_reader(BufReader::new(workbook_xml));
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) | Event::Empty(e) => {
                    if e.name().as_ref() == b"sheet" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"name" {
                                names.push(attr.unescape_value()?.to_string());
                            }
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

    fn parse_sheet_xml(
        &mut self,
        path: &str,
    ) -> Result<(
        HashMap<(u32, u32), Cell>,
        Vec<u32>,
        Vec<u32>,
        Vec<(u32, u32, u32, u32)>,
        usize,
        Vec<String>,
        Option<(u32, u32)>,
    )> {
        let mut cells = HashMap::new();
        let mut hidden_columns = Vec::new();
        let mut hidden_rows = Vec::new();
        let mut merged_cells = Vec::new();
        let mut cf_count = 0;
        let mut cf_ranges = Vec::new();
        let mut dim_range = None;
        let mut shared_formulas: HashMap<
            u32,
            Vec<(String, u32, u32, Option<(u32, u32, u32, u32)>)>,
        > = HashMap::new();

        let sheet_xml = self.archive.by_name(path)?;
        let mut reader = Reader::from_reader(BufReader::new(sheet_xml));
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_row = 0u32;
        let mut current_col = 0u32;

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => match e.name().as_ref() {
                    b"dimension" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ref" {
                                let ref_str = attr.unescape_value()?;
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
                    b"col" => {
                        let mut min = 0u32;
                        let mut max = 0u32;
                        let mut hidden = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"min" => {
                                    min = attr.unescape_value()?.parse::<u32>()?.saturating_sub(1);
                                }
                                b"max" => {
                                    max = attr.unescape_value()?.parse::<u32>()?.saturating_sub(1);
                                }
                                b"hidden" => {
                                    hidden = attr.value.as_ref() == b"1"
                                        || attr.value.as_ref() == b"true";
                                }
                                _ => {}
                            }
                        }
                        if hidden {
                            for col in min..=max {
                                hidden_columns.push(col);
                            }
                        }
                    }
                    b"row" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r" {
                                current_row =
                                    attr.unescape_value()?.parse::<u32>()?.saturating_sub(1);
                            }
                            if attr.key.as_ref() == b"hidden"
                                && (attr.value.as_ref() == b"1" || attr.value.as_ref() == b"true")
                            {
                                hidden_rows.push(current_row);
                            }
                        }
                        current_col = 0;
                    }
                    b"c" => {
                        let mut r_attr = String::new();
                        let mut s_attr = None;
                        let mut t_attr = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"r" => r_attr = attr.unescape_value()?.to_string(),
                                b"s" => s_attr = Some(attr.unescape_value()?.parse::<usize>()?),
                                b"t" => t_attr = attr.unescape_value()?.to_string(),
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

                        let (value, mut formula, shared_si, shared_ref) = parse_cell_contents(
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

                                if let Some((base_formula, base_row, base_col, _)) = best_def {
                                    let row_shift = row as i32 - *base_row as i32;
                                    let col_shift = col as i32 - *base_col as i32;
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
                            num_fmt,
                        };
                        if let Some(mut f) = formula {
                            if f.starts_with('=') {
                                f = f[1..].to_string();
                            }
                            cell.value = match cell.value {
                                CellValue::Formula {
                                    cached_error: Some(err),
                                    ..
                                } => CellValue::formula_with_error(f, err),
                                _ => CellValue::formula(f),
                            };
                        }
                        cells.insert((row, col), cell);
                    }
                    b"mergeCell" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ref" {
                                let ref_str = attr.unescape_value()?;
                                if let Some(range) = parse_cell_range(&ref_str) {
                                    merged_cells.push(range);
                                }
                            }
                        }
                    }
                    b"cfRule" => {
                        cf_count += 1;
                    }
                    b"conditionalFormatting" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"sqref" {
                                let sqref = attr.unescape_value()?;
                                cf_ranges.push(sqref.to_string());
                            }
                        }
                    }

                    _ => {}
                },
                Event::Empty(e) => match e.name().as_ref() {
                    b"dimension" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ref" {
                                let ref_str = attr.unescape_value()?;
                                if let Some((_, _, end_row, end_col)) = parse_cell_range(&ref_str) {
                                    dim_range = Some((end_row + 1, end_col + 1));
                                } else if let Some((end_row, end_col)) = parse_cell_ref(&ref_str) {
                                    dim_range = Some((end_row + 1, end_col + 1));
                                }
                            }
                        }
                    }
                    b"col" => {
                        let mut min = 0u32;
                        let mut max = 0u32;
                        let mut hidden = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"min" => {
                                    min = attr.unescape_value()?.parse::<u32>()?.saturating_sub(1);
                                }
                                b"max" => {
                                    max = attr.unescape_value()?.parse::<u32>()?.saturating_sub(1);
                                }
                                b"hidden" => {
                                    hidden = attr.value.as_ref() == b"1"
                                        || attr.value.as_ref() == b"true";
                                }
                                _ => {}
                            }
                        }
                        if hidden {
                            for col in min..=max {
                                hidden_columns.push(col);
                            }
                        }
                    }
                    b"c" => {
                        let mut r_attr = String::new();
                        let mut s_attr = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"r" => r_attr = attr.unescape_value()?.to_string(),
                                b"s" => s_attr = Some(attr.unescape_value()?.parse::<usize>()?),
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
                                num_fmt,
                            },
                        );
                    }
                    b"row" => {
                        // Empty row tag
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r" {
                                current_row =
                                    attr.unescape_value()?.parse::<u32>()?.saturating_sub(1);
                            }
                        }
                        current_col = 0;
                    }
                    b"mergeCell" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ref" {
                                let ref_str = attr.unescape_value()?;
                                if let Some(range) = parse_cell_range(&ref_str) {
                                    merged_cells.push(range);
                                }
                            }
                        }
                    }
                    b"cfRule" => {
                        cf_count += 1;
                    }
                    b"conditionalFormatting" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"sqref" {
                                let sqref = attr.unescape_value()?;
                                // Strict normalization reverted to preserve original format (e.g. A1 instead of A1:A1):
                                cf_ranges.push(sqref.into_owned());
                            }
                        }
                    }
                    _ => {}
                },
                Event::End(e) => {
                    if e.name().as_ref() == b"worksheet" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        // Include hidden rows/columns in used_range for format parity
        // Both ODS and XLSX should report ALL empty rows/columns (visible or hidden) if they have properties
        if let Some((mut rows, mut cols)) = dim_range {
            if let Some(&max_hidden_row) = hidden_rows.iter().max() {
                rows = rows.max(max_hidden_row + 1);
            }
            if let Some(&max_hidden_col) = hidden_columns.iter().max() {
                cols = cols.max(max_hidden_col + 1);
            }
            dim_range = Some((rows, cols));
        } else if !hidden_rows.is_empty() || !hidden_columns.is_empty() {
            let mut rows = 1u32;
            let mut cols = 1u32;
            if let Some(&max_hidden_row) = hidden_rows.iter().max() {
                rows = rows.max(max_hidden_row + 1);
            }
            if let Some(&max_hidden_col) = hidden_columns.iter().max() {
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

fn parse_cell_contents<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    t_attr: &str,
    shared_strings: &[String],
    _styles: &[String],
    num_fmt: Option<&str>,
) -> Result<(CellValue, Option<String>, Option<u32>, Option<String>)> {
    let mut value = CellValue::Empty;
    let mut formula = None;
    let mut shared_si = None;
    let mut shared_ref = None;
    let mut potential_error = None; // Store potential error value from t="e"
    let mut buf = Vec::new();

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => match e.name().as_ref() {
                b"v" => {
                    let v_text = if let Event::Start(_) = event {
                        read_text_node(reader)?
                    } else {
                        String::new()
                    };
                    value = match t_attr {
                        "s" => {
                            let idx = v_text.parse::<usize>().unwrap_or(0);
                            CellValue::Text(shared_strings.get(idx).cloned().unwrap_or_default())
                        }
                        "b" => CellValue::Boolean(v_text == "1"),
                        "e" => {
                            // Store the error value but don't create error cell yet
                            // We need to check if there's a formula first
                            potential_error = Some(v_text);
                            CellValue::Empty // Temporary, will be set later if needed
                        }
                        _ => {
                            // Check if this is a text-formatted number
                            // In XLSX, text format is indicated by num_fmt == "@"
                            if num_fmt == Some("@") {
                                // Store as text even if it looks like a number
                                CellValue::Text(v_text)
                            } else if let Ok(n) = v_text.parse::<f64>() {
                                CellValue::Number(n)
                            } else {
                                CellValue::Text(v_text)
                            }
                        }
                    };
                }
                b"f" => {
                    let mut si = None;
                    let mut is_shared = false;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"si" => {
                                si = attr.unescape_value()?.parse::<u32>().ok();
                            }
                            b"t" => {
                                if attr.value.as_ref() == b"shared" {
                                    is_shared = true;
                                }
                            }
                            b"ref" => {
                                shared_ref = Some(attr.unescape_value()?.to_string());
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
                b"is" => {
                    // Inline string can have multiple <t> tags
                    if let Event::Start(_) = event {
                        let mut is_text = String::new();
                        let mut is_buf = Vec::new();
                        loop {
                            match reader.read_event_into(&mut is_buf)? {
                                Event::Start(ref ee) if ee.name().as_ref() == b"t" => {
                                    is_text.push_str(&read_text_node(reader)?);
                                }
                                Event::End(ref ee) if ee.name().as_ref() == b"is" => break,
                                Event::Eof => break,
                                _ => {}
                            }
                            is_buf.clear();
                        }
                        value = CellValue::Text(is_text);
                    }
                }
                _ => {}
            },
            Event::End(e) if e.name().as_ref() == b"c" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    // Post-processing: handle potential_error
    // Excel uses t="e" for TWO different purposes:
    // 1. Real error cells (formula evaluated to error) - e.g., =car() -> #NAME?
    // 2. Array formulas (returns array, not error) - e.g., =Sheet!A1:B10 or =A1:A10
    //
    // The key distinction: array formulas contain range references (with ':')
    // Real errors are formulas without ranges that failed evaluation
    //
    // Note: In valid spreadsheet files, t="e" ALWAYS has a formula. Errors without
    // formulas don't exist in Excel/ODS files.
    if let Some(err) = potential_error
        && let Some(ref f) = formula
    {
        // If t="e" is present with a formula, it's a formula that evaluated to an error
        value = CellValue::formula_with_error(f.clone(), err);
    }

    Ok((value, formula, shared_si, shared_ref))
}

pub fn extract_shared_strings(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Vec<String>> {
    let mut strings = Vec::new();
    let ss_xml = match archive.by_name("xl/sharedStrings.xml") {
        Ok(file) => file,
        Err(_) => return Ok(strings),
    };

    let mut reader = Reader::from_reader(BufReader::new(ss_xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut current_string = String::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name().as_ref() == b"t" => {
                current_string.push_str(&read_text_node(&mut reader)?);
            }
            Event::End(e) if e.name().as_ref() == b"si" => {
                strings.push(current_string.clone());
                current_string.clear();
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(strings)
}

fn read_text_node<R: std::io::BufRead>(reader: &mut Reader<R>) -> Result<String> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Text(e) => text.push_str(e.unescape()?.as_ref()),
            Event::CData(e) => text.push_str(&String::from_utf8_lossy(e.as_ref())),
            Event::End(_) => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}
/// Count conditional formatting rules in a worksheet
pub fn count_conditional_formatting(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    sheet_name: &str,
) -> Result<usize> {
    let sheet_path = match get_xlsx_sheet_path(archive, sheet_name) {
        Ok(path) => path,
        Err(_) => return Ok(0),
    };

    let sheet_xml = match archive.by_name(&sheet_path) {
        Ok(file) => file,
        Err(_) => return Ok(0), // No sheet file, return 0
    };

    let buf_reader = BufReader::new(sheet_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut cf_count = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"cfRule" {
                    cf_count += 1;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(anyhow::anyhow!("XML parsing error: {}", e));
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
    let workbook_xml = match archive.by_name("xl/workbook.xml") {
        Ok(file) => file,
        Err(_) => return Ok(hidden_sheets), // No workbook.xml, return empty
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
                    b"sheets" => in_sheets = true,
                    b"sheet" if in_sheets => {
                        let mut name = String::new();
                        let mut state = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"name" => {
                                    name = attr.unescape_value()?.to_string();
                                }
                                b"state" => {
                                    state = attr.unescape_value()?.to_string();
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
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"sheets" {
                    in_sheets = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML parsing error: {}", e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(hidden_sheets)
}

/// Extract hidden columns and rows from a worksheet
pub fn extract_hidden_columns_rows_from_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    sheet_name: &str,
) -> Result<(Vec<u32>, Vec<u32>)> {
    let mut hidden_columns = Vec::new();
    let mut hidden_rows = Vec::new();

    let sheet_path = match get_xlsx_sheet_path(archive, sheet_name) {
        Ok(path) => path,
        Err(_) => return Ok((hidden_columns, hidden_rows)),
    };

    let sheet_xml = match archive.by_name(&sheet_path) {
        Ok(file) => file,
        Err(_) => return Ok((hidden_columns, hidden_rows)),
    };

    let buf_reader = BufReader::new(sheet_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                match e.name().as_ref() {
                    b"col" => {
                        // Column definition with hidden attribute
                        let mut min_col = 0u32;
                        let mut max_col = 0u32;
                        let mut hidden = false;

                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                match attr.key.as_ref() {
                                    b"min" => {
                                        if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                            min_col = val.saturating_sub(1); // Convert to 0-based
                                        }
                                    }
                                    b"max" => {
                                        if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                            max_col = val.saturating_sub(1); // Convert to 0-based
                                        }
                                    }
                                    b"hidden" => {
                                        hidden = attr.unescape_value()? == "1"
                                            || attr.unescape_value()?.to_lowercase() == "true";
                                    }
                                    _ => {}
                                }
                            }
                        }

                        if hidden {
                            for col in min_col..=max_col {
                                hidden_columns.push(col);
                            }
                        }
                    }
                    b"row" => {
                        // Row definition with hidden attribute
                        let mut row_num = 0u32;
                        let mut hidden = false;

                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                match attr.key.as_ref() {
                                    b"r" => {
                                        if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                            row_num = val.saturating_sub(1); // Convert to 0-based
                                        }
                                    }
                                    b"hidden" => {
                                        hidden = attr.unescape_value()? == "1"
                                            || attr.unescape_value()?.to_lowercase() == "true";
                                    }
                                    _ => {}
                                }
                            }
                        }

                        if hidden {
                            hidden_rows.push(row_num);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML parsing error: {}", e)),
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

    let sheet_path = match get_xlsx_sheet_path(archive, sheet_name) {
        Ok(path) => path,
        Err(_) => return Ok(merged_cells),
    };

    let sheet_xml = match archive.by_name(&sheet_path) {
        Ok(file) => file,
        Err(_) => return Ok(merged_cells),
    };

    let buf_reader = BufReader::new(sheet_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"mergeCell" {
                    for attr in e.attributes() {
                        if let Ok(attr) = attr
                            && attr.key.as_ref() == b"ref"
                        {
                            let ref_str = attr.unescape_value()?;
                            if let Some((start_row, start_col, end_row, end_col)) =
                                parse_cell_range(&ref_str)
                            {
                                merged_cells.push((start_row, start_col, end_row, end_col));
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML parsing error: {}", e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(merged_cells)
}

/// Parse a cell range like "A1:B2" into (start_row, start_col, end_row, end_col)
fn parse_cell_range(range: &str) -> Option<(u32, u32, u32, u32)> {
    let parts: Vec<&str> = range.split(':').collect();
    if parts.len() != 2 {
        return None;
    }

    let (start_row, start_col) = parse_cell_ref(parts[0])?;
    let (end_row, end_col) = parse_cell_ref(parts[1])?;

    Some((start_row, start_col, end_row, end_col))
}

/// Parse a cell reference like "A1" into (row, col) as 0-based indices
fn parse_cell_ref(cell_ref: &str) -> Option<(u32, u32)> {
    let mut col = 0u32;
    let mut row_str = String::new();

    for ch in cell_ref.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else if ch.is_ascii_digit() {
            row_str.push(ch);
        }
    }

    if row_str.is_empty() {
        return None;
    }

    let row = row_str.parse::<u32>().ok()?;

    // Convert to 0-based
    Some((row.saturating_sub(1), col.saturating_sub(1)))
}

/// Extract cell formats from XLSX file
/// Returns a list of format strings indexed by style index (xf index)
pub fn extract_formats_from_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Vec<String>> {
    let formats = Vec::new();
    let mut num_fmts = HashMap::new();

    // Built-in formats (simplified subset)
    // See https://github.com/dtjohnson/xlsx-populate/blob/master/lib/NumFmt.js
    num_fmts.insert(0, "General".to_string());
    num_fmts.insert(1, "0".to_string());
    num_fmts.insert(2, "0.00".to_string());
    num_fmts.insert(3, "#,##0".to_string());
    num_fmts.insert(4, "#,##0.00".to_string());
    num_fmts.insert(9, "0%".to_string());
    num_fmts.insert(10, "0.00%".to_string());
    num_fmts.insert(11, "0.00E+00".to_string());
    num_fmts.insert(12, "# ?/?".to_string());
    num_fmts.insert(13, "# ??/??".to_string());
    num_fmts.insert(14, "mm-dd-yy".to_string());
    num_fmts.insert(15, "d-mmm-yy".to_string());
    num_fmts.insert(16, "d-mmm".to_string());
    num_fmts.insert(17, "mmm-yy".to_string());
    num_fmts.insert(18, "h:mm AM/PM".to_string());
    num_fmts.insert(19, "h:mm:ss AM/PM".to_string());
    num_fmts.insert(20, "h:mm".to_string());
    num_fmts.insert(21, "h:mm:ss".to_string());
    num_fmts.insert(22, "m/d/yy h:mm".to_string());
    num_fmts.insert(37, "#,##0 ;(#,##0)".to_string());
    num_fmts.insert(38, "#,##0 ;[Red](#,##0)".to_string());
    num_fmts.insert(39, "#,##0.00;(#,##0.00)".to_string());
    num_fmts.insert(40, "#,##0.00;[Red](#,##0.00)".to_string());
    num_fmts.insert(45, "mm:ss".to_string());
    num_fmts.insert(46, "[h]:mm:ss".to_string());
    num_fmts.insert(47, "mmss.0".to_string());
    num_fmts.insert(48, "##0.0E+0".to_string());
    num_fmts.insert(49, "@".to_string());

    // Parse styles.xml
    let styles_xml = match archive.by_name("xl/styles.xml") {
        Ok(file) => file,
        Err(_) => return Ok(formats), // No styles, return empty
    };

    let buf_reader = BufReader::new(styles_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                match e.name().as_ref() {
                    b"numFmt" => {
                        let mut id = 0u32;
                        let mut code = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"numFmtId" => {
                                    if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                        id = val;
                                    }
                                }
                                b"formatCode" => {
                                    code =
                                        attr.unescape_value().unwrap_or_default().replace('\\', "");
                                }
                                _ => {}
                            }
                        }
                        if !code.is_empty() {
                            num_fmts.insert(id, code);
                        }
                    }
                    b"xf" => {
                        // We need to know if we are in cellXfs.
                        // However, quick-xml is stream. styleSheet usually has <cellStyleXfs> then <cellXfs>.
                        // We only care about <cellXfs>.
                        // Determining if we are in cellXfs is tricky without state.
                        // But usually cellXfs come after numFmts.
                        // And cellStyleXfs also use `xf` tag.
                        // We might need to track depth or parent.
                        // Or simplistic: gather all xfs and hope cellXfs are processed in order?
                        // No, `cellStyleXfs` are different.
                        // I'll track `in_cell_xfs`.
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }

    // Reset reader to parse xfs properly
    // ... Actually easier to do it in one pass if we track state.
    // Let's rewrite the loop logic above.
    Ok(Vec::new()) // Placeholder
}

/// Actual implementation of parsing styles
pub fn parse_styles(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Vec<String>> {
    let mut num_fmts = HashMap::new();

    // Built-in formats (same as above)
    num_fmts.insert(0, "General".to_string());
    num_fmts.insert(1, "0".to_string());
    num_fmts.insert(2, "0.00".to_string());
    num_fmts.insert(3, "#,##0".to_string());
    num_fmts.insert(4, "#,##0.00".to_string());
    num_fmts.insert(9, "0%".to_string());
    num_fmts.insert(10, "0.00%".to_string());
    num_fmts.insert(11, "0.00E+00".to_string());
    num_fmts.insert(12, "# ?/?".to_string());
    num_fmts.insert(13, "# ??/??".to_string());
    num_fmts.insert(14, "mm-dd-yy".to_string());
    num_fmts.insert(15, "d-mmm-yy".to_string());
    num_fmts.insert(16, "d-mmm".to_string());
    num_fmts.insert(17, "mmm-yy".to_string());
    num_fmts.insert(18, "h:mm AM/PM".to_string());
    num_fmts.insert(19, "h:mm:ss AM/PM".to_string());
    num_fmts.insert(20, "h:mm".to_string());
    num_fmts.insert(21, "h:mm:ss".to_string());
    num_fmts.insert(22, "m/d/yy h:mm".to_string());
    num_fmts.insert(37, "#,##0 ;(#,##0)".to_string());
    num_fmts.insert(38, "#,##0 ;[Red](#,##0)".to_string());
    num_fmts.insert(39, "#,##0.00;(#,##0.00)".to_string());
    num_fmts.insert(40, "#,##0.00;[Red](#,##0.00)".to_string());
    num_fmts.insert(45, "mm:ss".to_string());
    num_fmts.insert(46, "[h]:mm:ss".to_string());
    num_fmts.insert(47, "mmss.0".to_string());
    num_fmts.insert(48, "##0.0E+0".to_string());
    num_fmts.insert(49, "@".to_string());

    let styles_xml = match archive.by_name("xl/styles.xml") {
        Ok(file) => file,
        Err(_) => return Ok(Vec::new()),
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
                    b"numFmt" => {
                        let mut id = 0u32;
                        let mut code = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"numFmtId" => {
                                    if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                        id = val;
                                    }
                                }
                                b"formatCode" => {
                                    code =
                                        attr.unescape_value().unwrap_or_default().replace('\\', "");
                                }
                                _ => {}
                            }
                        }
                        if !code.is_empty() {
                            num_fmts.insert(id, code);
                        }
                    }
                    b"cellXfs" => {
                        in_cell_xfs = true;
                    }
                    b"xf" if in_cell_xfs => {
                        // Extract numFmtId
                        let mut num_fmt_id = 0u32;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr
                                && attr.key.as_ref() == b"numFmtId"
                                && let Ok(val) = attr.unescape_value()?.parse::<u32>()
                            {
                                num_fmt_id = val;
                            }
                        }
                        // Look up format code
                        let format_code = num_fmts
                            .get(&num_fmt_id)
                            .cloned()
                            .unwrap_or_else(|| "General".to_string());
                        xfs.push(format_code);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                match e.name().as_ref() {
                    b"numFmt" => {
                        // Same as above, handle empty tag if it occurs (unlikely for numFmt with attrs)
                        let mut id = 0u32;
                        let mut code = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"numFmtId" => {
                                    if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                        id = val;
                                    }
                                }
                                b"formatCode" => {
                                    code =
                                        attr.unescape_value().unwrap_or_default().replace('\\', "");
                                }
                                _ => {}
                            }
                        }
                        if !code.is_empty() {
                            num_fmts.insert(id, code);
                        }
                    }
                    b"xf" if in_cell_xfs => {
                        let mut num_fmt_id = 0u32;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr
                                && attr.key.as_ref() == b"numFmtId"
                                && let Ok(val) = attr.unescape_value()?.parse::<u32>()
                            {
                                num_fmt_id = val;
                            }
                        }
                        let format_code = num_fmts
                            .get(&num_fmt_id)
                            .cloned()
                            .unwrap_or_else(|| "General".to_string());
                        xfs.push(format_code);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"cellXfs" {
                    in_cell_xfs = false;
                }
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

    let sheet_path = match get_xlsx_sheet_path(archive, sheet_name) {
        Ok(path) => path,
        Err(_) => return Ok(cell_styles),
    };
    let sheet_xml = match archive.by_name(&sheet_path) {
        Ok(file) => file,
        Err(_) => return Ok(cell_styles),
    };

    let buf_reader = BufReader::new(sheet_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"c" {
                    // Cell element `c`
                    let mut row = 0u32;
                    let mut col = 0u32;
                    let mut style_index = 0usize;
                    let mut has_style = false;

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"r" => {
                                let r_str = attr.unescape_value()?;
                                if let Some((r, c)) = parse_cell_ref(&r_str) {
                                    row = r;
                                    col = c;
                                }
                            }
                            b"s" => {
                                if let Ok(val) = attr.unescape_value()?.parse::<usize>() {
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

    let sheet_path = match get_xlsx_sheet_path(archive, sheet_name) {
        Ok(path) => path,
        Err(_) => return Ok(formulas),
    };

    let sheet_xml = archive.by_name(&sheet_path)?;
    let mut reader = Reader::from_reader(BufReader::new(sheet_xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut current_cell_ref = String::new();
    let mut in_formula = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) => match e.name().as_ref() {
                b"c" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"r" {
                            current_cell_ref = attr.unescape_value()?.to_string();
                        }
                    }
                }
                b"f" => {
                    in_formula = true;
                    let mut si = None;
                    let mut t = None;

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"si" => {
                                si = attr.unescape_value()?.parse::<u32>().ok();
                            }
                            b"t" => {
                                t = Some(attr.unescape_value()?.to_string());
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
                        formulas.insert((r, c), format!("__SHARED__{}", s_idx));
                    }
                }
                _ => {}
            },
            Event::Text(e) if in_formula => {
                let formula_text = e.unescape()?.to_string();
                if let Some((r, c)) = parse_cell_ref(&current_cell_ref) {
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
            Event::End(e) => {
                if e.name().as_ref() == b"f" {
                    in_formula = false;
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
    use std::io::Write;

    #[test]
    fn test_extract_tables_from_xlsx() {
        use std::io::Cursor;
        use zip::write::FileOptions;

        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));

            let options =
                FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);

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
    assert!(sheets.len() > 0, "Should have at least one sheet");

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

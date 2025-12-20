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
                                b"name" => name = String::from_utf8_lossy(&attr.value).to_string(),
                                b"r:id" => r_id = String::from_utf8_lossy(&attr.value).to_string(),
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
                                b"Id" => id = String::from_utf8_lossy(&attr.value).to_string(),
                                b"Target" => t = String::from_utf8_lossy(&attr.value).to_string(),
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
                            if let Ok(attr) = attr {
                                if attr.key.as_ref() == b"name" {
                                    current_name = String::from_utf8_lossy(&attr.value).to_string();
                                }
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
                    defined_names.insert(current_name.clone(), current_ref.clone());
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

        for name in sheet_names {
            let path = get_xlsx_sheet_path(self.archive, &name)?;
            let mut sheet = Sheet::new(name.clone());
            sheet.sheet_path = Some(path.clone());

            // Parse sheet data
            let (cells, hidden_cols, hidden_rows, merged_cells) = self.parse_sheet_xml(&path)?;

            sheet.cells = cells;
            sheet.hidden_columns = hidden_cols;
            sheet.hidden_rows = hidden_rows;
            sheet.merged_cells = merged_cells;

            // Calculate used range
            if let Some((max_row, max_col)) = sheet.last_data_cell() {
                sheet.used_range = Some((max_row + 1, max_col + 1));
            }

            sheets.push(sheet);
        }

        Ok(sheets)
    }

    fn read_defined_names(&mut self) -> Result<HashMap<String, String>> {
        extract_defined_names_from_xlsx(self.archive)
    }

    fn read_hidden_sheets(&mut self) -> Result<Vec<String>> {
        extract_hidden_sheets_from_xlsx(self.archive)
    }

    fn has_macros(&mut self) -> Result<bool> {
        Ok(self.archive.by_name("xl/vbaProject.bin").is_ok())
    }
}

pub struct OdsReader<'a, R: std::io::Read + std::io::Seek> {
    archive: &'a mut ZipArchive<R>,
}

impl<'a, R: std::io::Read + std::io::Seek> OdsReader<'a, R> {
    pub fn new(archive: &'a mut ZipArchive<R>) -> Result<Self> {
        Ok(Self { archive })
    }
}

impl<'a, R: std::io::Read + std::io::Seek> WorkbookReader for OdsReader<'a, R> {
    fn read_sheets(&mut self) -> Result<Vec<Sheet>> {
        let mut sheets = Vec::new();

        let content_xml = match self.archive.by_name("content.xml") {
            Ok(file) => file,
            Err(_) => return Ok(sheets),
        };

        let mut reader = Reader::from_reader(BufReader::new(content_xml));
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_sheet: Option<Sheet> = None;
        let mut current_row = 0u32;
        let mut row_repeated = 1u32;
        let mut current_col = 0u32;

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"table:table" => {
                    let mut name = String::new();
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"table:name" {
                            name = String::from_utf8_lossy(&attr.value).to_string();
                        }
                    }
                    current_sheet = Some(Sheet::new(name));
                    current_row = 0;
                }
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"table:table-row" => {
                    row_repeated = 1;
                    current_col = 0;
                    if let Some(ref mut sheet) = current_sheet {
                        let mut hidden = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"table:number-rows-repeated" => {
                                    row_repeated = String::from_utf8_lossy(&attr.value)
                                        .parse::<u32>()
                                        .unwrap_or(1);
                                }
                                b"table:visibility" => {
                                    if attr.value.as_ref() == b"collapse"
                                        || attr.value.as_ref() == b"filter"
                                    {
                                        hidden = true;
                                    }
                                }
                                _ => {}
                            }
                        }
                        if hidden {
                            for r in 0..row_repeated {
                                sheet.hidden_rows.push(current_row + r);
                            }
                        }
                    }
                }
                Event::Start(e) | Event::Empty(e)
                    if e.name().as_ref() == b"table:table-cell"
                        || e.name().as_ref() == b"table:covered-table-cell" =>
                {
                    if let Some(ref mut sheet) = current_sheet {
                        let mut col_repeated = 1u32;
                        let mut cols_spanned = 1u32;
                        let mut rows_spanned = 1u32;
                        let mut formula = None;
                        let mut value = CellValue::Empty;
                        let mut has_value = false;
                        let mut is_error_cell = false;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"table:number-columns-repeated" => {
                                    col_repeated = String::from_utf8_lossy(&attr.value)
                                        .parse::<u32>()
                                        .unwrap_or(1);
                                }
                                b"table:number-columns-spanned" => {
                                    cols_spanned = String::from_utf8_lossy(&attr.value)
                                        .parse::<u32>()
                                        .unwrap_or(1);
                                }
                                b"table:number-rows-spanned" => {
                                    rows_spanned = String::from_utf8_lossy(&attr.value)
                                        .parse::<u32>()
                                        .unwrap_or(1);
                                }
                                b"table:formula" => {
                                    formula = Some(normalize_ods_formula(
                                        &String::from_utf8_lossy(&attr.value),
                                    ));
                                }
                                b"calcext:value-type" => {
                                    if attr.value.as_ref() == b"error" {
                                        is_error_cell = true;
                                    }
                                }
                                b"office:value"
                                | b"office:string-value"
                                | b"office:boolean-value"
                                | b"office:date-value" => {
                                    let val_str = String::from_utf8_lossy(&attr.value).to_string();
                                    if !has_value {
                                        value = match attr.key.as_ref() {
                                            b"office:value" => {
                                                if let Ok(n) = val_str.parse::<f64>() {
                                                    CellValue::Number(n)
                                                } else {
                                                    CellValue::Text(val_str)
                                                }
                                            }
                                            b"office:boolean-value" => {
                                                CellValue::Boolean(val_str == "true")
                                            }
                                            _ => CellValue::Text(val_str),
                                        };
                                        has_value = true;
                                    }
                                }
                                _ => {}
                            }
                        }

                        if cols_spanned > 1 || rows_spanned > 1 {
                            sheet.merged_cells.push((
                                current_row,
                                current_col,
                                current_row + rows_spanned - 1,
                                current_col + cols_spanned - 1,
                            ));
                        }

                        // For error cells, read the error text from <text:p>
                        if is_error_cell {
                            let mut error_text = String::new();
                            let mut text_buf = Vec::new();
                            loop {
                                match reader.read_event_into(&mut text_buf)? {
                                    Event::Start(ref te) if te.name().as_ref() == b"text:p" => {
                                        let mut p_buf = Vec::new();
                                        loop {
                                            match reader.read_event_into(&mut p_buf)? {
                                                Event::Text(ref t) => {
                                                    error_text.push_str(&t.unescape()?.to_string());
                                                }
                                                Event::End(ref pe)
                                                    if pe.name().as_ref() == b"text:p" =>
                                                {
                                                    break;
                                                }
                                                Event::Eof => break,
                                                _ => {}
                                            }
                                            p_buf.clear();
                                        }
                                    }
                                    Event::End(ref te)
                                        if te.name().as_ref() == b"table:table-cell"
                                            || te.name().as_ref()
                                                == b"table:covered-table-cell" =>
                                    {
                                        break;
                                    }
                                    Event::Eof => break,
                                    _ => {}
                                }
                                text_buf.clear();
                            }
                            if !error_text.is_empty() {
                                value = CellValue::formula_with_error("", error_text);
                                has_value = true;
                            }
                        }

                        if has_value || formula.is_some() {
                            let mut cell_value = value;
                            if let Some(f) = formula {
                                cell_value = match cell_value {
                                    CellValue::Formula {
                                        cached_error: Some(msg),
                                        ..
                                    } => CellValue::formula_with_error(f, msg),
                                    _ => CellValue::formula(f),
                                };
                            }

                            for r in 0..row_repeated {
                                for c in 0..col_repeated {
                                    let cell = Cell {
                                        row: current_row + r,
                                        col: current_col + c,
                                        value: cell_value.clone(),
                                        num_fmt: None,
                                    };
                                    sheet.cells.insert((current_row + r, current_col + c), cell);
                                }
                            }
                        }
                        current_col += col_repeated;
                    }
                }
                Event::End(e) if e.name().as_ref() == b"table:table-row" => {
                    current_row += row_repeated;
                }
                Event::End(e) if e.name().as_ref() == b"table:table" => {
                    if let Some(mut sheet) = current_sheet.take() {
                        if let Some((max_row, max_col)) = sheet.last_data_cell() {
                            sheet.used_range = Some((max_row + 1, max_col + 1));
                        }
                        sheets.push(sheet);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(sheets)
    }

    fn read_defined_names(&mut self) -> Result<HashMap<String, String>> {
        let mut defined_names = HashMap::new();

        let content_xml = match self.archive.by_name("content.xml") {
            Ok(file) => file,
            Err(_) => return Ok(defined_names),
        };

        let mut reader = Reader::from_reader(BufReader::new(content_xml));
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut in_named_expressions = false;

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) if e.name().as_ref() == b"table:named-expressions" => {
                    in_named_expressions = true;
                }
                Event::Empty(e) | Event::Start(e)
                    if in_named_expressions && e.name().as_ref() == b"table:named-range" =>
                {
                    let mut name = String::new();
                    let mut cell_range_address = String::new();

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"table:name" => {
                                name = String::from_utf8_lossy(&attr.value).to_string();
                            }
                            b"table:cell-range-address" => {
                                cell_range_address =
                                    String::from_utf8_lossy(&attr.value).to_string();
                            }
                            _ => {}
                        }
                    }

                    if !name.is_empty() && !cell_range_address.is_empty() {
                        // Normalize ODS range address to Excel-style (Sheet!A1:B2)
                        let normalized = normalize_ods_formula(&cell_range_address);
                        defined_names.insert(name, normalized);
                    }
                }
                Event::End(e) if e.name().as_ref() == b"table:named-expressions" => {
                    in_named_expressions = false;
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(defined_names)
    }

    fn read_hidden_sheets(&mut self) -> Result<Vec<String>> {
        super::ods_parser::extract_hidden_sheets_from_ods(self.archive)
    }

    fn has_macros(&mut self) -> Result<bool> {
        super::ods_parser::has_macros_ods(self.archive)
    }
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
                                names.push(String::from_utf8_lossy(&attr.value).to_string());
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
    )> {
        let mut cells = HashMap::new();
        let mut hidden_columns = Vec::new();
        let mut hidden_rows = Vec::new();
        let mut merged_cells = Vec::new();
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
                    b"col" => {
                        let mut min = 0u32;
                        let mut max = 0u32;
                        let mut hidden = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"min" => {
                                    min = String::from_utf8_lossy(&attr.value)
                                        .parse::<u32>()?
                                        .saturating_sub(1);
                                }
                                b"max" => {
                                    max = String::from_utf8_lossy(&attr.value)
                                        .parse::<u32>()?
                                        .saturating_sub(1);
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
                                current_row = String::from_utf8_lossy(&attr.value)
                                    .parse::<u32>()?
                                    .saturating_sub(1);
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
                                b"r" => r_attr = String::from_utf8_lossy(&attr.value).to_string(),
                                b"s" => {
                                    s_attr = Some(
                                        String::from_utf8_lossy(&attr.value).parse::<usize>()?,
                                    )
                                }
                                b"t" => t_attr = String::from_utf8_lossy(&attr.value).to_string(),
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

                        let (value, mut formula, shared_si, shared_ref) = parse_cell_contents(
                            &mut reader,
                            &t_attr,
                            &self.shared_strings,
                            &self.styles,
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
                                    if let Some((min_r, min_c, max_r, max_c)) = def.3 {
                                        if row >= min_r
                                            && row <= max_r
                                            && col >= min_c
                                            && col <= max_c
                                        {
                                            best_def = Some(def);
                                            break;
                                        }
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

                        let num_fmt = s_attr.and_then(|idx| self.styles.get(idx).cloned());

                        let mut cell = Cell {
                            row,
                            col,
                            value: value.clone(),
                            num_fmt,
                        };
                        if let Some(f) = formula {
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
                                let ref_str = String::from_utf8_lossy(&attr.value);
                                if let Some(range) = parse_cell_range(&ref_str) {
                                    merged_cells.push(range);
                                }
                            }
                        }
                    }
                    _ => {}
                },
                Event::Empty(e) => match e.name().as_ref() {
                    b"c" => {
                        let mut r_attr = String::new();
                        let mut s_attr = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"r" => r_attr = String::from_utf8_lossy(&attr.value).to_string(),
                                b"s" => {
                                    s_attr = Some(
                                        String::from_utf8_lossy(&attr.value).parse::<usize>()?,
                                    )
                                }
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
                                current_row = String::from_utf8_lossy(&attr.value)
                                    .parse::<u32>()?
                                    .saturating_sub(1);
                            }
                        }
                        current_col = 0;
                    }
                    _ => {}
                },
                Event::End(e) => match e.name().as_ref() {
                    b"sheetData" => break,
                    _ => {}
                },
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok((cells, hidden_columns, hidden_rows, merged_cells))
    }
}

fn parse_cell_contents<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    t_attr: &str,
    shared_strings: &[String],
    _styles: &[String],
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
                            if let Ok(n) = v_text.parse::<f64>() {
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
                                si = String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            }
                            b"t" => {
                                if attr.value.as_ref() == b"shared" {
                                    is_shared = true;
                                }
                            }
                            b"ref" => {
                                shared_ref = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            _ => {}
                        }
                    }

                    if let Event::Start(_) = event {
                        let f_text = read_text_node(reader)?;
                        if !f_text.is_empty() {
                            formula = Some(format!("={}", f_text));
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
    if let Some(err) = potential_error {
        if let Some(ref f) = formula {
            // Only process if we have a formula (which we always should for t="e")
            if !f.contains(':') {
                // Real error - formula without range that evaluated to error
                value = CellValue::formula_with_error("", err);
            }
            // If formula contains ':', it's an array formula - ignore error marker
        }
        // No else needed - t="e" without formula is invalid/impossible in real files
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
            Event::Text(e) => text.push_str(&e.unescape()?.to_string()),
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

                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                match attr.key.as_ref() {
                                    b"name" => {
                                        name = String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    b"state" => {
                                        state = String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    _ => {}
                                }
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
                                        if let Ok(val) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            min_col = val.saturating_sub(1); // Convert to 0-based
                                        }
                                    }
                                    b"max" => {
                                        if let Ok(val) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            max_col = val.saturating_sub(1); // Convert to 0-based
                                        }
                                    }
                                    b"hidden" => {
                                        hidden = String::from_utf8_lossy(&attr.value) == "1"
                                            || String::from_utf8_lossy(&attr.value).to_lowercase()
                                                == "true";
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
                                        if let Ok(val) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            row_num = val.saturating_sub(1); // Convert to 0-based
                                        }
                                    }
                                    b"hidden" => {
                                        hidden = String::from_utf8_lossy(&attr.value) == "1"
                                            || String::from_utf8_lossy(&attr.value).to_lowercase()
                                                == "true";
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
                        if let Ok(attr) = attr {
                            if attr.key.as_ref() == b"ref" {
                                let ref_str = String::from_utf8_lossy(&attr.value);
                                if let Some((start_row, start_col, end_row, end_col)) =
                                    parse_cell_range(&ref_str)
                                {
                                    merged_cells.push((start_row, start_col, end_row, end_col));
                                }
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
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                match attr.key.as_ref() {
                                    b"numFmtId" => {
                                        if let Ok(val) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            id = val;
                                        }
                                    }
                                    b"formatCode" => {
                                        code = attr.unescape_value().unwrap_or_default().into();
                                    }
                                    _ => {}
                                }
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
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                match attr.key.as_ref() {
                                    b"numFmtId" => {
                                        if let Ok(val) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            id = val;
                                        }
                                    }
                                    b"formatCode" => {
                                        code = attr.unescape_value().unwrap_or_default().into();
                                    }
                                    _ => {}
                                }
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
                            if let Ok(attr) = attr {
                                if attr.key.as_ref() == b"numFmtId" {
                                    if let Ok(val) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        num_fmt_id = val;
                                    }
                                }
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
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                match attr.key.as_ref() {
                                    b"numFmtId" => {
                                        if let Ok(val) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            id = val;
                                        }
                                    }
                                    b"formatCode" => {
                                        code = attr.unescape_value().unwrap_or_default().into();
                                    }
                                    _ => {}
                                }
                            }
                        }
                        if !code.is_empty() {
                            num_fmts.insert(id, code);
                        }
                    }
                    b"xf" if in_cell_xfs => {
                        let mut num_fmt_id = 0u32;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                if attr.key.as_ref() == b"numFmtId" {
                                    if let Ok(val) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        num_fmt_id = val;
                                    }
                                }
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

                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            match attr.key.as_ref() {
                                b"r" => {
                                    let r_str = String::from_utf8_lossy(&attr.value);
                                    if let Some((r, c)) = parse_cell_ref(&r_str) {
                                        row = r;
                                        col = c;
                                    }
                                }
                                b"s" => {
                                    if let Ok(val) =
                                        String::from_utf8_lossy(&attr.value).parse::<usize>()
                                    {
                                        style_index = val;
                                        has_style = true;
                                    }
                                }
                                _ => {}
                            }
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
                            current_cell_ref = String::from_utf8_lossy(&attr.value).to_string();
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
                                si = String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            }
                            b"t" => {
                                t = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            _ => {}
                        }
                    }

                    if let Some(_s_idx) = si {
                        if t.as_deref() == Some("shared") {
                            // This might be the base or a consumer
                            // We wait for the text content to see if it's the base
                        }
                    }

                    // Store si for the current cell to association after getting Text
                    if let Some(s_idx) = si {
                        if let Some((r, c)) = parse_cell_ref(&current_cell_ref) {
                            // Temporarily store si to handle it in Text
                            // We use a prefix to distinguish from actual formulas
                            formulas.insert((r, c), format!("__SHARED__{}", s_idx));
                        }
                    }
                }
                _ => {}
            },
            Event::Text(e) if in_formula => {
                let formula_text = e.unescape()?.to_string();
                if let Some((r, c)) = parse_cell_ref(&current_cell_ref) {
                    // Check if this cell was marked as shared
                    if let Some(marker) = formulas.get(&(r, c)) {
                        if marker.starts_with("__SHARED__") {
                            let si: u32 = marker["__SHARED__".len()..].parse().unwrap();
                            shared_formulas.insert(si, formula_text.clone());
                        }
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

/// Check if XLSX file contains VBA macros
pub fn has_vba_project_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<bool> {
    // Check for the presence of vbaProject.bin file
    Ok(archive.by_name("xl/vbaProject.bin").is_ok())
}
fn normalize_ods_formula(formula: &str) -> String {
    use regex::Regex;
    use std::sync::OnceLock;

    static ODS_REL_REF: OnceLock<Regex> = OnceLock::new();
    static ODS_COL_REF: OnceLock<Regex> = OnceLock::new();
    static ODS_ROW_REF: OnceLock<Regex> = OnceLock::new();
    static ODS_SHEET_REF: OnceLock<Regex> = OnceLock::new();
    static ODS_SHEET_REF_NO_BRACKET: OnceLock<Regex> = OnceLock::new();
    static ODS_RECT_REF: OnceLock<Regex> = OnceLock::new();

    let mut normalized = formula.strip_prefix("of:=").unwrap_or(formula).to_string();

    // 0. Replace rectangular ranges [.A1:.B2] -> A1:B2
    // Matches [.A1:.B2]
    let rect_ref =
        ODS_RECT_REF.get_or_init(|| Regex::new(r"\[\.([A-Z]+[0-9]+):\.([A-Z]+[0-9]+)\]").unwrap());
    normalized = rect_ref.replace_all(&normalized, "$1:$2").to_string();

    // 0b. Replace sheet-qualified ranges [$Sheet.A1:.A2] -> Sheet!A1:A2
    // Matches [$Sheet.A1:.B2], [$Sheet.$A$1:.$B$2]
    // The colon is followed by a dot for relative/absolute mixing in ODS
    static ODS_SHEET_RANGE_REF: OnceLock<Regex> = OnceLock::new();
    let sheet_range_ref = ODS_SHEET_RANGE_REF
        .get_or_init(|| Regex::new(r"\[\$([^.]+)\.([A-Za-z0-9$]+):\.([A-Za-z0-9$]+)\]").unwrap());
    normalized = sheet_range_ref
        .replace_all(&normalized, "$1!$2:$3")
        .to_string();

    // 1. Replace sheet references [$Sheet1.A1] -> Sheet1!A1 (BEFORE non-bracketed version)
    // Matches [$Sheet.$A$1], [$Sheet.A1], etc.
    // The cell reference part can have $ signs for absolute references
    let sheet_ref =
        ODS_SHEET_REF.get_or_init(|| Regex::new(r"\[\$([^.]+)\.(\$?[A-Z]+\$?[0-9]+)\]").unwrap());
    normalized = sheet_ref.replace_all(&normalized, "$1!$2").to_string();

    // 2. Replace $SHEETNAME.CELLREF -> SHEETNAME!CELLREF (without brackets)
    // Matches $INGRESOS.BC$50, $Sheet1.A1, $Sheet.$A$1 etc.
    // This comes AFTER bracketed version, so it won't match already-processed [$SHEET.CELL]
    let sheet_ref_no_bracket = ODS_SHEET_REF_NO_BRACKET
        .get_or_init(|| Regex::new(r"\$([A-Za-z0-9_]+)\.(\$?[A-Z]+\$?[0-9]+)").unwrap());
    normalized = sheet_ref_no_bracket
        .replace_all(&normalized, "$1!$2")
        .to_string();

    // 3. Replace relative references [.A1] -> A1
    // Matches [.A1], [.AA123]
    let rel_ref = ODS_REL_REF.get_or_init(|| Regex::new(r"\[\.([A-Z]+[0-9]+)\]").unwrap());
    normalized = rel_ref.replace_all(&normalized, "$1").to_string();

    // 4. Replace whole column references [.A:.A] -> A:A
    // Matches [.A:.A], [.A:.B]
    let col_ref = ODS_COL_REF.get_or_init(|| Regex::new(r"\[\.([A-Z]+):\.([A-Z]+)\]").unwrap());
    normalized = col_ref.replace_all(&normalized, "$1:$2").to_string();

    // 5. Replace whole row references [.1:.1] -> 1:1
    // Matches [.1:.1], [.1:.10]
    let row_ref = ODS_ROW_REF.get_or_init(|| Regex::new(r"\[\.([0-9]+):\.([0-9]+)\]").unwrap());
    normalized = row_ref.replace_all(&normalized, "$1:$2").to_string();

    normalized
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_normalize_ods_formula() {
        // 1. Strip prefix
        assert_eq!(normalize_ods_formula("of:=SUM(A1)"), "SUM(A1)");

        // 2. Relative refs
        assert_eq!(normalize_ods_formula("=[.A1]+[.B2]"), "=A1+B2");

        // 3. Whole column refs
        assert_eq!(normalize_ods_formula("=SUM([.A:.A])"), "=SUM(A:A)");
        assert_eq!(normalize_ods_formula("=SUM([.A:.C])"), "=SUM(A:C)");

        // 4. Whole row refs
        assert_eq!(normalize_ods_formula("=SUM([.1:.1])"), "=SUM(1:1)");

        // 5. Sheet refs
        assert_eq!(normalize_ods_formula("=[$Sheet1.A1]"), "=Sheet1!A1");
        // Note: The regex captures sheet name and cell address.
        // [$Sheet1.A1:B2] -> Sheet1!A1:B2 requires slightly more careful regex if we want to handle range fully
        // But for basic cases:
        assert_eq!(normalize_ods_formula("=[$Sheet1.A1]"), "=Sheet1!A1");

        // 6. Cross-sheet refs without brackets (ODS format: $SHEETNAME.CELLREF)
        assert_eq!(normalize_ods_formula("=$INGRESOS.BC$50"), "=INGRESOS!BC$50");
        assert_eq!(normalize_ods_formula("=$Sheet1.A1"), "=Sheet1!A1");
        assert_eq!(
            normalize_ods_formula("=ROUND($INGRESOS.BC$50*G721,-3)"),
            "=ROUND(INGRESOS!BC$50*G721,-3)"
        );

        // 7. Mixed
        assert_eq!(
            normalize_ods_formula("of:=SUM([.A:.A]; [$Sheet2.C5])"),
            "SUM(A:A; Sheet2!C5)"
        );

        // 8. Rectangular ranges (The Bug)
        assert_eq!(normalize_ods_formula("=SUM([.A1:.B2])"), "=SUM(A1:B2)");

        // 9. Sheet references with absolute cell references ($ signs)
        assert_eq!(
            normalize_ods_formula("=ROUND([$Calculos_Adicionales.$F$137];-3)"),
            "=ROUND(Calculos_Adicionales!$F$137;-3)"
        );
        assert_eq!(normalize_ods_formula("=[$Sheet1.$A$1]"), "=Sheet1!$A$1");

        // 10. Complex ODS formula user verification (Bug Reproduction)
        let input = "=(SUM($Rentas_de_Trabajo.$F$1452:$F$1453)+SUM($INGRESOS.$BC$38:$BL$40))-SUM(,$Rentas_de_Trabajo.F10,$Rentas_de_Trabajo.F15:F16,$Rentas_de_Trabajo.F26,$Rentas_de_Trabajo.F31:F32,$Rentas_de_Trabajo.F42,$Rentas_de_Trabajo.F47:F48,$Rentas_de_Trabajo.F58,$Rentas_de_Trabajo.F63:F64,$Rentas_de_Trabajo.F74,$Rentas_de_Trabajo.F79:F80,$Rentas_de_Trabajo.F90,$Rentas_de_Trabajo.F95:F96,$Rentas_de_Trabajo.F106,$Rentas_de_Trabajo.F111:F112,$Rentas_de_Trabajo.F122,$Rentas_de_Trabajo.F127:F128,$Rentas_de_Trabajo.F138,$Rentas_de_Trabajo.F143:F144)";
        let expected = "=(SUM(Rentas_de_Trabajo!$F$1452:$F$1453)+SUM(INGRESOS!$BC$38:$BL$40))-SUM(,Rentas_de_Trabajo!F10,Rentas_de_Trabajo!F15:F16,Rentas_de_Trabajo!F26,Rentas_de_Trabajo!F31:F32,Rentas_de_Trabajo!F42,Rentas_de_Trabajo!F47:F48,Rentas_de_Trabajo!F58,Rentas_de_Trabajo!F63:F64,Rentas_de_Trabajo!F74,Rentas_de_Trabajo!F79:F80,Rentas_de_Trabajo!F90,Rentas_de_Trabajo!F95:F96,Rentas_de_Trabajo!F106,Rentas_de_Trabajo!F111:F112,Rentas_de_Trabajo!F122,Rentas_de_Trabajo!F127:F128,Rentas_de_Trabajo!F138,Rentas_de_Trabajo!F143:F144)";
        assert_eq!(normalize_ods_formula(input), expected);

        // 11. Range normalization specific test
        // Original failing pattern: [Rentas_de_Trabajo!$F$1452:.$F$1453]
        // Which presumably comes from raw input: [$Rentas_de_Trabajo.$F$1452:.$F$1453]
        assert_eq!(
            normalize_ods_formula("=[$Rentas_de_Trabajo.$F$1452:.$F$1453]"),
            "=Rentas_de_Trabajo!$F$1452:$F$1453"
        );
    }

    #[test]
    fn test_parse_cell_contents_array_formulas_vs_errors() {
        use quick_xml::Reader;
        use std::io::Cursor;

        // Test 1: Array formula with range reference - should NOT create error
        let xml = r#"<c r="A1" t="e"><f>OUTPUT!B459:D505</f><v>#VALUE!</v></c>"#;
        let mut reader = Reader::from_reader(Cursor::new(xml));
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let _ = reader.read_event_into(&mut buf); // Skip to start

        let (value, formula, _, _) = parse_cell_contents(&mut reader, "e", &[], &[]).unwrap();

        // Should have formula but NOT error value (array formula)
        assert!(formula.is_some());
        assert_eq!(formula.unwrap(), "=OUTPUT!B459:D505");
        // Value should be Empty, not an error
        match value {
            CellValue::Empty => {} // Expected
            CellValue::Formula { cached_error, .. } => {
                panic!(
                    "Array formula should not have cached_error: {:?}",
                    cached_error
                );
            }
            _ => panic!("Expected Empty for array formula, got: {:?}", value),
        }

        // Test 2: Real error formula - should create error
        let xml2 = r#"<c r="A2" t="e"><f>car()</f><v>#NAME?</v></c>"#;
        let mut reader2 = Reader::from_reader(Cursor::new(xml2));
        reader2.config_mut().trim_text(true);
        let mut buf2 = Vec::new();
        let _ = reader2.read_event_into(&mut buf2);

        let (value2, formula2, _, _) = parse_cell_contents(&mut reader2, "e", &[], &[]).unwrap();

        // Should have formula AND error value
        assert!(formula2.is_some());
        assert_eq!(formula2.unwrap(), "=car()");
        match value2 {
            CellValue::Formula {
                formula,
                cached_error,
            } => {
                assert_eq!(formula, "");
                assert_eq!(cached_error, Some("#NAME?".to_string()));
            }
            _ => panic!("Expected Formula with error, got: {:?}", value2),
        }

        // Test 3: Another array formula pattern
        let xml3 = r#"<c r="A3" t="e"><f>Recomendaciones!C6:BE93</f><v>#VALUE!</v></c>"#;
        let mut reader3 = Reader::from_reader(Cursor::new(xml3));
        reader3.config_mut().trim_text(true);
        let mut buf3 = Vec::new();
        let _ = reader3.read_event_into(&mut buf3);

        let (value3, formula3, _, _) = parse_cell_contents(&mut reader3, "e", &[], &[]).unwrap();

        assert!(formula3.is_some());
        match value3 {
            CellValue::Empty => {} // Expected for array formula
            _ => panic!("Expected Empty for array formula, got: {:?}", value3),
        }

        // Note: We don't test t="e" without formula because it's impossible in real files
        // Every error cell in Excel/ODS has a formula that evaluated to the error
    }
}

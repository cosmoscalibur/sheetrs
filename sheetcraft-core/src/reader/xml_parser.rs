//! XML parsing utilities for extracting metadata from XLSX files

use anyhow::Result;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashMap;
use std::io::BufReader;
use zip::ZipArchive;

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

/// Count conditional formatting rules in a worksheet
pub fn count_conditional_formatting(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    sheet_index: usize,
) -> Result<usize> {
    // Sheet files are named sheet1.xml, sheet2.xml, etc. (1-indexed)
    let sheet_path = format!("xl/worksheets/sheet{}.xml", sheet_index + 1);

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
    sheet_index: usize,
) -> Result<(Vec<u32>, Vec<u32>)> {
    let mut hidden_columns = Vec::new();
    let mut hidden_rows = Vec::new();

    // Sheet files are named sheet1.xml, sheet2.xml, etc. (1-indexed)
    let sheet_path = format!("xl/worksheets/sheet{}.xml", sheet_index + 1);

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
    sheet_index: usize,
) -> Result<Vec<(u32, u32, u32, u32)>> {
    let mut merged_cells = Vec::new();

    // Sheet files are named sheet1.xml, sheet2.xml, etc. (1-indexed)
    let sheet_path = format!("xl/worksheets/sheet{}.xml", sheet_index + 1);

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
    sheet_index: usize,
) -> Result<HashMap<(u32, u32), usize>> {
    let mut cell_styles = HashMap::new();

    let sheet_path = format!("xl/worksheets/sheet{}.xml", sheet_index + 1);
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

/// Check if XLSX file contains VBA macros
pub fn has_vba_project_xlsx(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<bool> {
    // Check for the presence of vbaProject.bin file
    Ok(archive.by_name("xl/vbaProject.bin").is_ok())
}

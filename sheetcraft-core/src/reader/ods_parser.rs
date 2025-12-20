//! ODS parsing utilities for extracting metadata from ODS files

use anyhow::Result;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashMap;
use std::io::BufReader;
use zip::ZipArchive;

use super::{Cell, CellValue, Sheet, WorkbookReader};

/// Extract hidden sheets from ODS file
/// ODS format: <table:table table:name="SheetName" table:display="false">
pub fn extract_hidden_sheets_from_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Vec<String>> {
    let mut hidden_sheets = Vec::new();

    // Try to read content.xml
    let content_xml = match archive.by_name("content.xml") {
        Ok(file) => file,
        Err(_) => return Ok(hidden_sheets), // No content.xml, return empty
    };

    let buf_reader = BufReader::new(content_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"table:table" {
                    let mut name = String::new();
                    let mut display = String::from("true"); // Default is visible

                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            match attr.key.as_ref() {
                                b"table:name" => {
                                    name = String::from_utf8_lossy(&attr.value).to_string();
                                }
                                b"table:display" => {
                                    display = String::from_utf8_lossy(&attr.value).to_string();
                                }
                                _ => {}
                            }
                        }
                    }

                    // If display is "false", the sheet is hidden
                    if !name.is_empty() && display == "false" {
                        hidden_sheets.push(name);
                    }
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

/// Extract hidden columns and rows from an ODS worksheet
/// ODS format:
/// - Hidden columns: <table:table-column table:visibility="collapse" or "filter">
/// - Hidden rows: <table:table-row table:visibility="collapse" or "filter">
pub fn extract_hidden_columns_rows_from_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    sheet_name: &str,
) -> Result<(Vec<u32>, Vec<u32>)> {
    let mut hidden_columns = Vec::new();
    let mut hidden_rows = Vec::new();

    // ODS stores all sheets in content.xml
    let content_xml = match archive.by_name("content.xml") {
        Ok(file) => file,
        Err(_) => return Ok((hidden_columns, hidden_rows)),
    };

    let buf_reader = BufReader::new(content_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_target_sheet = false;
    let mut current_col = 0u32;
    let mut current_row = 0u32;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                match e.name().as_ref() {
                    b"table:table" => {
                        let mut name = String::new();
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"table:name" {
                                name = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }

                        if name == sheet_name {
                            in_target_sheet = true;
                            current_col = 0;
                            current_row = 0;
                        } else if in_target_sheet {
                            // We were in the target sheet, but now we encountered another table tag.
                            // In ODS, tables are top-level and don't nest.
                            break;
                        }
                    }
                    b"table:table-column" if in_target_sheet => {
                        let mut visibility = String::from("visible");
                        let mut repeated = 1u32;

                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                match attr.key.as_ref() {
                                    b"table:visibility" => {
                                        visibility =
                                            String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    b"table:number-columns-repeated" => {
                                        if let Ok(val) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            repeated = val;
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }

                        // If visibility is "collapse" or "filter", the column is hidden
                        if visibility == "collapse" || visibility == "filter" {
                            for _ in 0..repeated {
                                hidden_columns.push(current_col);
                                current_col += 1;
                            }
                        } else {
                            current_col += repeated;
                        }
                    }
                    b"table:table-row" if in_target_sheet => {
                        let mut visibility = String::from("visible");
                        let mut repeated = 1u32;

                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                match attr.key.as_ref() {
                                    b"table:visibility" => {
                                        visibility =
                                            String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    b"table:number-rows-repeated" => {
                                        if let Ok(val) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            repeated = val;
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }

                        // If visibility is "collapse" or "filter", the row is hidden
                        if visibility == "collapse" || visibility == "filter" {
                            for _ in 0..repeated {
                                hidden_rows.push(current_row);
                                current_row += 1;
                            }
                        } else {
                            current_row += repeated;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table" && in_target_sheet {
                    break; // Found our sheet, done
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

/// Extract merged cell ranges from an ODS worksheet
/// ODS format: <table:table-cell table:number-columns-spanned="X" table:number-rows-spanned="Y">
pub fn extract_merged_cells_from_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    sheet_name: &str,
) -> Result<Vec<(u32, u32, u32, u32)>> {
    let mut merged_cells = Vec::new();

    // ODS stores all sheets in content.xml
    let content_xml = match archive.by_name("content.xml") {
        Ok(file) => file,
        Err(_) => return Ok(merged_cells),
    };

    let buf_reader = BufReader::new(content_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_target_sheet = false;
    let mut current_row = 0u32;
    let mut current_col = 0u32;
    let mut current_row_repeated = 1u32;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                match e.name().as_ref() {
                    b"table:table" => {
                        let mut name = String::new();
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"table:name" {
                                name = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }

                        if name == sheet_name {
                            in_target_sheet = true;
                            current_row = 0;
                            current_col = 0;
                        } else if in_target_sheet {
                            break;
                        }
                    }
                    b"table:table-row" if in_target_sheet => {
                        current_col = 0; // Reset column for new row
                        // Tracking row repetition for spanning logic
                        current_row_repeated = 1;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"table:number-rows-repeated" {
                                if let Ok(val) = String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    current_row_repeated = val;
                                }
                            }
                        }
                    }
                    b"table:table-cell" | b"table:covered-table-cell" if in_target_sheet => {
                        let mut cols_spanned = 1u32;
                        let mut rows_spanned = 1u32;
                        let mut repeated = 1u32;

                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                match attr.key.as_ref() {
                                    b"table:number-columns-spanned" => {
                                        if let Ok(val) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            cols_spanned = val;
                                        }
                                    }
                                    b"table:number-rows-spanned" => {
                                        if let Ok(val) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            rows_spanned = val;
                                        }
                                    }
                                    b"table:number-columns-repeated" => {
                                        if let Ok(val) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            repeated = val;
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }

                        // If either span is > 1, this is a merged cell
                        if cols_spanned > 1 || rows_spanned > 1 {
                            merged_cells.push((
                                current_row,
                                current_col,
                                current_row + rows_spanned - 1,
                                current_col + cols_spanned - 1,
                            ));
                        }

                        current_col += repeated;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) if in_target_sheet => match e.name().as_ref() {
                b"table:table-row" => {
                    current_row += current_row_repeated;
                }
                b"table:table" => {
                    break;
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML parsing error: {}", e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(merged_cells)
}

/// Check if ODS file contains macros
/// ODS macros are stored in Basic/ or Scripts/ directories
pub fn has_macros_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<bool> {
    // Check for Basic/ directory (LibreOffice Basic macros)
    for i in 0..archive.len() {
        if let Ok(file) = archive.by_index(i) {
            let name = file.name();
            if name.starts_with("Basic/") || name.starts_with("Scripts/") {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

/// Extract cached error values from an ODS worksheet
/// ODS error values are often stored in calcext:value-type="error" and calcext:value="#ERROR!"
pub fn extract_cached_errors_from_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    sheet_name: &str,
) -> Result<std::collections::HashMap<(u32, u32), String>> {
    use std::collections::HashMap;
    let mut errors = HashMap::new();

    // ODS stores all sheets in content.xml
    let content_xml = match archive.by_name("content.xml") {
        Ok(file) => file,
        Err(_) => return Ok(errors),
    };

    let buf_reader = BufReader::new(content_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_target_sheet = false;
    let mut current_row = 0u32;
    let mut current_col = 0u32;
    let mut current_row_repeated = 1u32;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table:table" => {
                    let mut name = String::new();
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"table:name" {
                            name = String::from_utf8_lossy(&attr.value).to_string();
                        }
                    }

                    if name == sheet_name {
                        in_target_sheet = true;
                        current_row = 0;
                        current_col = 0;
                    } else if in_target_sheet {
                        break;
                    }
                }
                b"table:table-row" if in_target_sheet => {
                    current_col = 0;
                    current_row_repeated = 1;
                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            if attr.key.as_ref() == b"table:number-rows-repeated" {
                                if let Ok(val) = String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    current_row_repeated = val;
                                }
                            }
                        }
                    }
                }
                b"table:table-cell" if in_target_sheet => {
                    let mut repeated = 1u32;
                    let mut is_error = false;
                    let mut error_msg = String::new();

                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            match attr.key.as_ref() {
                                b"calcext:value-type" => {
                                    if attr.value.as_ref() == b"error" {
                                        is_error = true;
                                    }
                                }
                                b"calcext:value" | b"office:string-value" => {
                                    if error_msg.is_empty() {
                                        error_msg =
                                            String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                }
                                b"table:number-columns-repeated" => {
                                    if let Ok(val) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        repeated = val;
                                    }
                                }
                                _ => {}
                            }
                        }
                    }

                    if is_error && error_msg.is_empty() {
                        error_msg = "#ERROR!".to_string();
                    }

                    if is_error && !error_msg.is_empty() {
                        for i in 0..repeated {
                            for r in 0..current_row_repeated {
                                errors
                                    .insert((current_row + r, current_col + i), error_msg.clone());
                            }
                        }
                    }
                    current_col += repeated;
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-row" if in_target_sheet => {
                    let mut repeated = 1u32;
                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            if attr.key.as_ref() == b"table:number-rows-repeated" {
                                if let Ok(val) = String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    repeated = val;
                                }
                            }
                        }
                    }
                    current_row += repeated;
                }
                b"table:table-cell" if in_target_sheet => {
                    let mut repeated = 1u32;
                    let mut is_error = false;
                    let mut error_msg = String::new();

                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            match attr.key.as_ref() {
                                b"calcext:value-type" => {
                                    if attr.value.as_ref() == b"error" {
                                        is_error = true;
                                    }
                                }
                                b"calcext:value" | b"office:string-value" => {
                                    if error_msg.is_empty() {
                                        error_msg =
                                            String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                }
                                b"table:number-columns-repeated" => {
                                    if let Ok(val) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        repeated = val;
                                    }
                                }
                                _ => {}
                            }
                        }
                    }

                    if is_error && error_msg.is_empty() {
                        error_msg = "#ERROR!".to_string();
                    }

                    if is_error && !error_msg.is_empty() {
                        for i in 0..repeated {
                            for r in 0..current_row_repeated {
                                errors
                                    .insert((current_row + r, current_col + i), error_msg.clone());
                            }
                        }
                    }
                    current_col += repeated;
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                match e.name().as_ref() {
                    b"table:table-row" if in_target_sheet => {
                        current_row += current_row_repeated;
                    }
                    b"table:table" if in_target_sheet => {
                        break; // Found our sheet, done
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

    Ok(errors)
}

/// Extract formulas from an ODS worksheet
/// ODS formulas are stored in table:formula attribute
pub fn extract_formulas_from_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    sheet_name: &str,
) -> Result<std::collections::HashMap<(u32, u32), String>> {
    use std::collections::HashMap;
    let mut formulas = HashMap::new();

    // ODS stores all sheets in content.xml
    let content_xml = match archive.by_name("content.xml") {
        Ok(file) => file,
        Err(_) => return Ok(formulas),
    };

    let buf_reader = BufReader::new(content_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_target_sheet = false;
    let mut current_row = 0u32;
    let mut current_col = 0u32;
    let mut current_row_repeated = 1u32;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table:table" => {
                    let mut name = String::new();
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"table:name" {
                            name = String::from_utf8_lossy(&attr.value).to_string();
                        }
                    }

                    if name == sheet_name {
                        in_target_sheet = true;
                        current_row = 0;
                        current_col = 0;
                    } else if in_target_sheet {
                        break;
                    }
                }
                b"table:table-row" if in_target_sheet => {
                    current_col = 0;
                    current_row_repeated = 1;
                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            if attr.key.as_ref() == b"table:number-rows-repeated" {
                                if let Ok(val) = String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    current_row_repeated = val;
                                }
                            }
                        }
                    }
                }
                b"table:table-cell" | b"table:covered-table-cell" if in_target_sheet => {
                    let mut repeated = 1u32;
                    let mut formula = String::new();

                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            match attr.key.as_ref() {
                                b"table:formula" => {
                                    formula = String::from_utf8_lossy(&attr.value).to_string();
                                }
                                b"table:number-columns-repeated" => {
                                    if let Ok(val) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        repeated = val;
                                    }
                                }
                                _ => {}
                            }
                        }
                    }

                    // Only extract formulas from non-repeated cells.
                    // Repeated cells should be handled by calamine, which correctly adjusts
                    // cell references. Copying the same formula to repeated cells without
                    // adjustment causes false circular references (e.g., G721's formula
                    // copied to G731 without adjusting references).
                    if !formula.is_empty() && repeated == 1 && current_row_repeated == 1 {
                        // ODS formulas usually start with "of:=". We keep it raw here and normalize later.
                        formulas.insert((current_row, current_col), formula);
                    }
                    current_col += repeated;
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-row" if in_target_sheet => {
                    let mut repeated = 1u32;
                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            if attr.key.as_ref() == b"table:number-rows-repeated" {
                                if let Ok(val) = String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    repeated = val;
                                }
                            }
                        }
                    }
                    current_row += repeated;
                }
                b"table:table-cell" | b"table:covered-table-cell" if in_target_sheet => {
                    let mut repeated = 1u32;
                    let mut formula = String::new();

                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            match attr.key.as_ref() {
                                b"table:formula" => {
                                    formula = String::from_utf8_lossy(&attr.value).to_string();
                                }
                                b"table:number-columns-repeated" => {
                                    if let Ok(val) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        repeated = val;
                                    }
                                }
                                _ => {}
                            }
                        }
                    }

                    // Only extract formulas from non-repeated cells (see comment above)
                    if !formula.is_empty() && repeated == 1 && current_row_repeated == 1 {
                        formulas.insert((current_row, current_col), formula);
                    }
                    current_col += repeated;
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                match e.name().as_ref() {
                    b"table:table-row" if in_target_sheet => {
                        current_row += current_row_repeated;
                    }
                    b"table:table" if in_target_sheet => {
                        break; // Found our sheet, done
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

    Ok(formulas)
}
pub fn normalize_ods_formula(formula: &str) -> String {
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
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"table:table-column" => {
                    if let Some(ref mut sheet) = current_sheet {
                        let mut hidden = false;
                        let mut repeated = 1u32;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"table:visibility" => {
                                    if attr.value.as_ref() == b"collapse"
                                        || attr.value.as_ref() == b"filter"
                                    {
                                        hidden = true;
                                    }
                                }
                                b"table:number-columns-repeated" => {
                                    repeated = String::from_utf8_lossy(&attr.value)
                                        .parse::<u32>()
                                        .unwrap_or(1);
                                }
                                _ => {}
                            }
                        }
                        if hidden {
                            for _ in 0..repeated {
                                sheet.hidden_columns.push(current_col);
                                current_col += 1;
                            }
                        } else {
                            current_col += repeated;
                        }
                    }
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

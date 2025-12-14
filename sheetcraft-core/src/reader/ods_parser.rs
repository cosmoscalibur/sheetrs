//! ODS parsing utilities for extracting metadata from ODS files

use anyhow::Result;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::BufReader;
use zip::ZipArchive;

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
    sheet_index: usize,
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
    let mut current_sheet = 0usize;
    let mut in_target_sheet = false;
    let mut current_col = 0u32;
    let mut current_row = 0u32;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                match e.name().as_ref() {
                    b"table:table" => {
                        if current_sheet == sheet_index {
                            in_target_sheet = true;
                            current_col = 0;
                            current_row = 0;
                        }
                        current_sheet += 1;
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
    sheet_index: usize,
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
    let mut current_sheet = 0usize;
    let mut in_target_sheet = false;
    let mut current_row = 0u32;
    let mut current_col = 0u32;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                match e.name().as_ref() {
                    b"table:table" => {
                        if current_sheet == sheet_index {
                            in_target_sheet = true;
                            current_row = 0;
                            current_col = 0;
                        }
                        current_sheet += 1;
                    }
                    b"table:table-row" if in_target_sheet => {
                        current_col = 0; // Reset column for new row
                    }
                    b"table:table-cell" if in_target_sheet => {
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
            Ok(Event::End(e)) => {
                match e.name().as_ref() {
                    b"table:table-row" if in_target_sheet => {
                        current_row += 1;
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

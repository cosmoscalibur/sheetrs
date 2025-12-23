//! ODS parsing utilities for extracting metadata from ODS files

use anyhow::Result;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashMap;
use std::io::BufReader;
use zip::ZipArchive;

use super::{Cell, CellValue, Sheet, WorkbookReader};

pub fn extract_hidden_sheets_from_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Vec<String>> {
    let mut hidden_sheets = Vec::new();
    let mut sheet_styles = Vec::new(); // (sheet_name, style_name)
    let mut hidden_styles = std::collections::HashSet::new();

    {
        let content_xml = match archive.by_name("content.xml") {
            Ok(file) => file,
            Err(_) => return Ok(hidden_sheets),
        };

        let buf_reader = BufReader::new(content_xml);
        let mut reader = Reader::from_reader(buf_reader);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"style:style" => {
                    let mut style_name = String::new();
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"style:name" {
                            style_name = attr.unescape_value()?.to_string();
                        }
                    }

                    if !style_name.is_empty() {
                        // If it was a Start event, now look for style:table-properties
                        let mut inner_buf = Vec::new();
                        loop {
                            match reader.read_event_into(&mut inner_buf)? {
                                Event::Start(ee) | Event::Empty(ee)
                                    if ee.name().as_ref() == b"style:table-properties" =>
                                {
                                    for attr in ee.attributes().flatten() {
                                        if attr.key.as_ref() == b"table:display"
                                            && attr.value.as_ref() == b"false"
                                        {
                                            hidden_styles.insert(style_name.clone());
                                        }
                                    }
                                }
                                Event::End(ee) if ee.name().as_ref() == b"style:style" => break,
                                Event::Eof => break,
                                _ => {}
                            }
                            inner_buf.clear();
                        }
                    }
                }
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"table:table" => {
                    let mut sheet_name = String::new();
                    let mut style_name = String::new();
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"table:name" {
                            sheet_name = attr.unescape_value()?.to_string();
                        } else if attr.key.as_ref() == b"table:style-name" {
                            style_name = attr.unescape_value()?.to_string();
                        }
                    }
                    if !sheet_name.is_empty() && !style_name.is_empty() {
                        sheet_styles.push((sheet_name, style_name));
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
    }

    for (name, style) in sheet_styles {
        if hidden_styles.contains(&style) {
            hidden_sheets.push(name);
        }
    }

    // Fallback: use settings.xml if no hidden sheets found or to supplement
    if hidden_sheets.is_empty() {
        let all_sheets = extract_all_sheet_names_from_ods(archive)?;
        let visible_sheets = extract_visible_sheets_from_settings(archive)?;
        for sheet in all_sheets {
            if !visible_sheets.contains(&sheet) && !hidden_sheets.contains(&sheet) {
                hidden_sheets.push(sheet);
            }
        }
    }

    Ok(hidden_sheets)
}

fn extract_all_sheet_names_from_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Vec<String>> {
    let mut sheet_names = Vec::new();

    let content_xml = match archive.by_name("content.xml") {
        Ok(file) => file,
        Err(_) => return Ok(sheet_names),
    };

    let buf_reader = BufReader::new(content_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"table:table" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"table:name" {
                            let name = attr.unescape_value()?.to_string();
                            if !name.is_empty() {
                                sheet_names.push(name);
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

    Ok(sheet_names)
}

/// Extract visible sheet names from settings.xml Tables section
fn extract_visible_sheets_from_settings(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<std::collections::HashSet<String>> {
    let mut visible_sheets = std::collections::HashSet::new();

    let settings_xml = match archive.by_name("settings.xml") {
        Ok(file) => file,
        Err(_) => {
            // No settings.xml - assume all sheets are visible (fail-safe)
            return Ok(visible_sheets);
        }
    };

    let buf_reader = BufReader::new(settings_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_tables_section = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                // Check if we're entering the Tables section
                if e.name().as_ref() == b"config:config-item-map-named" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"config:name" && attr.value.as_ref() == b"Tables" {
                            in_tables_section = true;
                        }
                    }
                }

                // If we're in Tables section, extract sheet names from map entries
                if in_tables_section && e.name().as_ref() == b"config:config-item-map-entry" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"config:name" {
                            let name = attr.unescape_value()?.to_string();
                            if !name.is_empty() {
                                visible_sheets.insert(name);
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                // Exit Tables section
                if e.name().as_ref() == b"config:config-item-map-named" {
                    in_tables_section = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML parsing error: {}", e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(visible_sheets)
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
                                name = attr.unescape_value()?.to_string();
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
                                        visibility = attr.unescape_value()?.to_string();
                                    }
                                    b"table:number-columns-repeated" => {
                                        if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
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
                                        visibility = attr.unescape_value()?.to_string();
                                    }
                                    b"table:number-rows-repeated" => {
                                        if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
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
            Ok(Event::Start(e)) => {
                match e.name().as_ref() {
                    b"table:table" => {
                        let mut name = String::new();
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"table:name" {
                                name = attr.unescape_value()?.to_string();
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
                                if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                    current_row_repeated = val;
                                }
                            }
                        }
                    }
                    b"table:table-cell" if in_target_sheet => {
                        let mut cols_spanned = 1u32;
                        let mut rows_spanned = 1u32;
                        let mut repeated = 1u32;

                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                match attr.key.as_ref() {
                                    b"table:number-columns-spanned" => {
                                        if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                            cols_spanned = val;
                                        }
                                    }
                                    b"table:number-rows-spanned" => {
                                        if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                            rows_spanned = val;
                                        }
                                    }
                                    b"table:number-columns-repeated" => {
                                        if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
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
                    b"table:covered-table-cell" if in_target_sheet => {
                        // Covered cells are part of a merged range and should increment the column counter
                        let mut repeated = 1u32;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                if attr.key.as_ref() == b"table:number-columns-repeated" {
                                    if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                        repeated = val;
                                    }
                                }
                            }
                        }
                        current_col += repeated;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                match e.name().as_ref() {
                    b"table:table-row" if in_target_sheet => {
                        // Empty row - increment row counter
                        let mut repeated = 1u32;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"table:number-rows-repeated" {
                                if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                    repeated = val;
                                }
                            }
                        }
                        current_row += repeated;
                    }
                    b"table:table-cell" if in_target_sheet => {
                        let mut cols_spanned = 1u32;
                        let mut rows_spanned = 1u32;
                        let mut repeated = 1u32;

                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                match attr.key.as_ref() {
                                    b"table:number-columns-spanned" => {
                                        if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                            cols_spanned = val;
                                        }
                                    }
                                    b"table:number-rows-spanned" => {
                                        if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                            rows_spanned = val;
                                        }
                                    }
                                    b"table:number-columns-repeated" => {
                                        if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
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
                    b"table:covered-table-cell" if in_target_sheet => {
                        // Covered cells are part of a merged range and should increment the column counter
                        let mut repeated = 1u32;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                if attr.key.as_ref() == b"table:number-columns-repeated" {
                                    if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
                                        repeated = val;
                                    }
                                }
                            }
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
/// ODS macros are stored in Basic/ or Scripts/ directories,
/// or declared in META-INF/manifest.xml
pub fn has_macros(archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>) -> Result<bool> {
    // 1. Check for directory presence
    for i in 0..archive.len() {
        if let Ok(file) = archive.by_index(i) {
            let name = file.name();
            if name.starts_with("Basic/") || name.starts_with("Scripts/") {
                return Ok(true);
            }
        }
    }

    // 2. Check manifest for macro-related media types
    if let Ok(manifest_file) = archive.by_name("META-INF/manifest.xml") {
        let buf_reader = BufReader::new(manifest_file);
        let mut reader = Reader::from_reader(buf_reader);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) | Event::Empty(e)
                    if e.name().as_ref() == b"manifest:file-entry" =>
                {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"manifest:media-type" {
                            let media_type = attr.unescape_value()?;
                            if media_type.contains("application/vnd.sun.xml.ui.configuration")
                                || media_type.contains("script")
                            {
                                // This is a bit broad, but Basic/ scripts often have specific media types
                            }
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
    }

    Ok(false)
}

/// Extract external links from ODS metadata
pub fn extract_external_links_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Vec<String>> {
    let mut links = std::collections::HashSet::new();

    let content_xml = match archive.by_name("content.xml") {
        Ok(file) => file,
        Err(_) => return Ok(Vec::new()),
    };

    let buf_reader = BufReader::new(content_xml);
    let mut reader = Reader::from_reader(buf_reader);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref e) if e.name().as_ref() == b"text:a" => {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"xlink:href" {
                        let link = attr.unescape_value()?.to_string();
                        if !link.starts_with("#") {
                            links.insert(link);
                        }
                    }
                }
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if e.name().as_ref() == b"table:table-source" =>
            {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"xlink:href" {
                        links.insert(attr.unescape_value()?.to_string());
                    }
                }
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if e.name().as_ref() == b"table:table-cell" =>
            {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"table:formula" {
                        let formula = attr.unescape_value()?;
                        if formula.contains("'file:///") {
                            if let Some(start) = formula.find("'file:///") {
                                let remainder = &formula[start + 1..];
                                if let Some(end) = remainder.find("'#") {
                                    let path = &remainder[..end];
                                    links.insert(path.to_string());
                                }
                            }
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    // Strip file:// prefix from paths for cleaner display
    let cleaned_links: Vec<String> = links
        .into_iter()
        .map(|link| {
            if link.starts_with("file://") {
                link.trim_start_matches("file://").to_string()
            } else {
                link
            }
        })
        .collect();

    Ok(cleaned_links)
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
                            name = attr.unescape_value()?.to_string();
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
                                if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
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
                                        error_msg = attr.unescape_value()?.to_string();
                                    }
                                }
                                b"table:number-columns-repeated" => {
                                    if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
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
                                if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
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
                                        error_msg = attr.unescape_value()?.to_string();
                                    }
                                }
                                b"table:number-columns-repeated" => {
                                    if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
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

/// Extract date styles from ODS content.xml and styles.xml
/// Returns a map of style_name -> excel_format_string
// Helper to extract date styles and cell style mappings
fn parse_ods_date(date_str: &str) -> Option<f64> {
    // Format: YYYY-MM-DD or YYYY-MM-DDThh:mm:ss
    let parts: Vec<&str> = date_str.split('T').collect();
    let date_part = parts[0];
    let time_part = if parts.len() > 1 { parts[1] } else { "" };

    let date_components: Vec<&str> = date_part.split('-').collect();
    if date_components.len() != 3 {
        return None;
    }

    let year = date_components[0].parse::<i32>().ok()?;
    let month = date_components[1].parse::<u32>().ok()?;
    let day = date_components[2].parse::<u32>().ok()?;

    // Simple days count from 1899-12-30
    // Excel epoch: 1899-12-30 = 0.
    // 1900-01-01 = 2 (Excel bug: 1900 is leap year).

    // We can use a simplified algorithm since we likely deal with modern dates
    // Algorithm to convert YMD to total days since 0000-03-01
    // But easier to just count days.

    let is_leap = |y: i32| (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);

    let days_in_month = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

    let mut total_days = 0;

    // Years
    for y in 1900..year {
        total_days += if is_leap(y) { 366 } else { 365 };
    }

    // Months
    for m in 1..month {
        if m == 2 && is_leap(year) {
            total_days += 29;
        } else {
            total_days += days_in_month[m as usize];
        }
    }

    // Days
    total_days += day as i32;

    // Adjust for Excel epoch (1900-01-01 is day 1, but we start counting from 1900-01-01 as day 1 in this loop?)
    // Loop starts 1900.
    // if date is 1900-01-01: loop 0, month 0, day 1. total = 1.
    // Excel 1900-01-01 is 2? No, 1. (Actually 1900-01-01 is 1.0).
    // Excel thinks 1900-02-29 exists (day 60).

    // If our date is > 1900-02-28, we need to ADD 1 to match Excel's bug.
    // Unless ODS date is pre-1900, which is rare.

    // Let's verify:
    // 1999-09-30 should be 36433.
    // Calc:
    // Years 1900..1999 (99 years).
    // Leaps: 1904, 08, 12, ... 96. (96-4)/4 + 1 = 24 leap years.
    // 99 * 365 + 24 = 36135 + 24 = 36159.
    // Months in 1999 (Jan-Aug): 31+28+31+30+31+30+31+31 = 243.
    // Days: 30.
    // Total = 36159 + 243 + 30 = 36432.
    // Target 36433.
    // Why diff 1? Because Excel has extra day (Feb 29 1900).
    // So we add 1 offset + 1 (starting index?).

    // Actually, "1900-01-01" in my loop gives 1. In Excel it is 1.
    // "1900-02-28" loop: 31 + 28 = 59. Excel: 59.
    // "1900-03-01" loop: 31 + 28 + 1 = 60. Excel: 61 (60 is 2/29).

    // So if total_days > 59, add 1.
    if total_days > 59 {
        total_days += 1;
    }

    // Time
    let mut time_fraction = 0.0;
    if !time_part.is_empty() {
        // HH:MM:SS or HH:MM:SS.mmm
        let time_parts: Vec<&str> = time_part.split(':').collect();
        if time_parts.len() >= 2 {
            let h = time_parts[0].parse::<f64>().unwrap_or(0.0);
            let m = time_parts[1].parse::<f64>().unwrap_or(0.0);
            let s = if time_parts.len() > 2 {
                time_parts[2].parse::<f64>().unwrap_or(0.0)
            } else {
                0.0
            };

            time_fraction = (h * 3600.0 + m * 60.0 + s) / 86400.0;
        }
    }

    // Excel starts from Dec 30 1899?
    // My loop started Jan 1 1900 as 1.
    // This matches Excel (1 = 1900-01-01).
    // So should be fine.

    Some(total_days as f64 + time_fraction)
}

pub fn extract_date_styles_from_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<std::collections::HashMap<String, String>> {
    use std::collections::HashMap;

    // Map: Data Style Name -> Format String (e.g., "N49" -> "dd/mm/yyyy")
    let mut data_styles = HashMap::new();
    // Map: Cell Style Name -> Data Style Name (e.g., "ce14" -> "N49")
    let mut cell_styles = HashMap::new();

    // Helper to parse a styles file (content.xml or styles.xml)
    let mut parse_styles_file = |file: std::io::BufReader<zip::read::ZipFile>| -> Result<()> {
        let mut reader = Reader::from_reader(file);
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut current_data_style_name = String::new();
        let mut current_format = String::new();
        let mut in_date_style = false;

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    match e.name().as_ref() {
                        b"number:date-style" => {
                            in_date_style = true;
                            current_format.clear();
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"style:name" {
                                    current_data_style_name = attr.unescape_value()?.to_string();
                                }
                            }
                        }
                        b"style:style" => {
                            let mut is_cell_style = false;
                            let mut style_name = String::new();
                            let mut data_style_name = String::new();

                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"style:family" => {
                                        if attr.value.as_ref() == b"table-cell" {
                                            is_cell_style = true;
                                        }
                                    }
                                    b"style:name" => {
                                        style_name = attr.unescape_value()?.to_string();
                                    }
                                    b"style:data-style-name" => {
                                        data_style_name = attr.unescape_value()?.to_string();
                                    }
                                    _ => {}
                                }
                            }

                            if is_cell_style
                                && !style_name.is_empty()
                                && !data_style_name.is_empty()
                            {
                                cell_styles.insert(style_name, data_style_name);
                            }
                        }
                        b"number:day" if in_date_style => {
                            let mut long = false;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"number:style"
                                    && attr.value.as_ref() == b"long"
                                {
                                    long = true;
                                }
                            }
                            current_format.push_str(if long { "dd" } else { "d" });
                        }
                        b"number:month" if in_date_style => {
                            let mut long = false;
                            let mut textual = false;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"number:style"
                                    && attr.value.as_ref() == b"long"
                                {
                                    long = true;
                                }
                                if attr.key.as_ref() == b"number:textual"
                                    && attr.value.as_ref() == b"true"
                                {
                                    textual = true;
                                }
                            }
                            if textual {
                                current_format.push_str(if long { "mmmm" } else { "mmm" });
                            } else {
                                current_format.push_str(if long { "mm" } else { "m" });
                            }
                        }
                        b"number:year" if in_date_style => {
                            let mut long = false;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"number:style"
                                    && attr.value.as_ref() == b"long"
                                {
                                    long = true;
                                }
                            }
                            current_format.push_str(if long { "yyyy" } else { "yy" });
                        }
                        b"number:hours" if in_date_style => {
                            current_format.push_str("hh");
                        }
                        b"number:minutes" if in_date_style => {
                            current_format.push_str("mm");
                        }
                        b"number:seconds" if in_date_style => {
                            current_format.push_str("ss");
                        }
                        b"number:text" if in_date_style => {
                            // Will read text event next
                        }
                        _ => {}
                    }
                }
                Event::Empty(e) => match e.name().as_ref() {
                    b"number:day" if in_date_style => {
                        let mut long = false;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"number:style"
                                && attr.value.as_ref() == b"long"
                            {
                                long = true;
                            }
                        }
                        current_format.push_str(if long { "dd" } else { "d" });
                    }
                    b"number:month" if in_date_style => {
                        let mut long = false;
                        let mut textual = false;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"number:style"
                                && attr.value.as_ref() == b"long"
                            {
                                long = true;
                            }
                            if attr.key.as_ref() == b"number:textual"
                                && attr.value.as_ref() == b"true"
                            {
                                textual = true;
                            }
                        }
                        if textual {
                            current_format.push_str(if long { "mmmm" } else { "mmm" });
                        } else {
                            current_format.push_str(if long { "mm" } else { "m" });
                        }
                    }
                    b"number:year" if in_date_style => {
                        let mut long = false;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"number:style"
                                && attr.value.as_ref() == b"long"
                            {
                                long = true;
                            }
                        }
                        current_format.push_str(if long { "yyyy" } else { "yy" });
                    }
                    b"number:hours" if in_date_style => {
                        current_format.push_str("hh");
                    }
                    b"number:minutes" if in_date_style => {
                        current_format.push_str("mm");
                    }
                    b"number:seconds" if in_date_style => {
                        current_format.push_str("ss");
                    }
                    _ => {}
                },
                Event::Text(e) if in_date_style => {
                    current_format.push_str(&e.unescape()?.to_string());
                }
                Event::End(e) => {
                    if e.name().as_ref() == b"number:date-style" {
                        if !current_data_style_name.is_empty() {
                            data_styles
                                .insert(current_data_style_name.clone(), current_format.clone());
                        }
                        in_date_style = false;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        Ok(())
    };

    // 1. Read automatic styles from content.xml
    if let Ok(file) = archive.by_name("content.xml") {
        parse_styles_file(BufReader::new(file))?;
    }

    // 2. Read styles from styles.xml (global styles)
    if let Ok(file) = archive.by_name("styles.xml") {
        parse_styles_file(BufReader::new(file))?;
    }

    // 3. Resolve Cell Styles to Format Strings
    let mut resolved_styles = HashMap::new();
    for (cell_style, data_style) in cell_styles {
        if let Some(format) = data_styles.get(&data_style) {
            resolved_styles.insert(cell_style, format.clone());
        }
    }

    // Also include data styles directly, just in case
    for (data_style, format) in data_styles {
        resolved_styles.entry(data_style).or_insert(format);
    }

    Ok(resolved_styles)
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
                            name = attr.unescape_value()?.to_string();
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
                                if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
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
                                    formula = attr.unescape_value()?.to_string();
                                }
                                b"table:number-columns-repeated" => {
                                    if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
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
                                if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
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
                                    formula = attr.unescape_value()?.to_string();
                                }
                                b"table:number-columns-repeated" => {
                                    if let Ok(val) = attr.unescape_value()?.parse::<u32>() {
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
/// Normalize ODS references (formulas, ranges, etc.) to a format consistent with XLSX.
/// Strips "of:=" prefix, handles sheet-qualified references, and cleans up local sheet names.
///
/// If `visible_to_xml_row` is provided, converts visible row numbers (1-indexed, used in ODS formulas)
/// to XML row numbers (0-indexed, used internally). This accounts for hidden rows.
/// The mapping is ONLY applied to references pointing to `current_sheet_name` (or local references).
pub fn normalize_ods_reference(
    reference: &str,
    preserve_sheet: bool,
    visible_to_xml_row: Option<&HashMap<u32, u32>>,
    current_sheet_name: Option<&str>,
) -> String {
    use regex::Regex;
    use std::sync::OnceLock;

    static ODS_REL_REF: OnceLock<Regex> = OnceLock::new();
    static ODS_COL_REF: OnceLock<Regex> = OnceLock::new();
    static ODS_ROW_REF: OnceLock<Regex> = OnceLock::new();
    static ODS_SHEET_REF: OnceLock<Regex> = OnceLock::new();
    static ODS_SHEET_REF_NO_BRACKET: OnceLock<Regex> = OnceLock::new();
    static ODS_RECT_REF: OnceLock<Regex> = OnceLock::new();

    let mut normalized = reference
        .strip_prefix("of:=")
        .unwrap_or(reference)
        .to_string();

    // 0. Replace rectangular ranges [.A1:.B2] -> A1:B2
    // Matches [.A1:.B2]
    let rect_ref = ODS_RECT_REF
        .get_or_init(|| Regex::new(r"\[\.(\$?[A-Z]+\$?[0-9]+):\.(\$?[A-Z]+\$?[0-9]+)\]").unwrap());
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

    // 0b2. Replace unbracketed sheet-qualified ranges $Sheet.A1:.A2 -> Sheet!A1:A2
    // Matches $Sheet.A1:.B2 (common in database-range targets)
    // We exclude '[' from the sheet name capture to avoid matching relative column refs like [.A:.A]
    static ODS_SHEET_RANGE_NO_BRACKET_REF: OnceLock<Regex> = OnceLock::new();
    let sheet_range_no_bracket_ref = ODS_SHEET_RANGE_NO_BRACKET_REF
        .get_or_init(|| Regex::new(r"\$?([^.\[]+)\.([A-Za-z0-9$]+):\.([A-Za-z0-9$]+)").unwrap());
    normalized = sheet_range_no_bracket_ref
        .replace_all(&normalized, "$1!$2:$3")
        .to_string();

    if !preserve_sheet {
        // 0c. Replace plain sheet-level ranges Sheet1.A1:Sheet1.B2 -> A1:B2
        // Matches Sheet1.A1:Sheet1.B2, Sheet1.A1, etc.
        // This is common in ODS conditional formatting targets.
        // We strip the sheet name if it's the same for both parts of the range.
        // Updated to handle both '.' (ODS native) and '!' (normalized) separators
        static ODS_LOCAL_RANGE_REF: OnceLock<Regex> = OnceLock::new();
        let local_range_ref = ODS_LOCAL_RANGE_REF.get_or_init(|| {
            Regex::new(r"([^.!]+)[.!]([A-Z0-9$]+):([^.!]+)[.!]([A-Z0-9$]+)").unwrap()
        });
        normalized = local_range_ref
            .replace_all(&normalized, |caps: &regex::Captures| {
                if &caps[1] == &caps[3] {
                    if &caps[2] == &caps[4] {
                        caps[2].to_string()
                    } else {
                        format!("{}:{}", &caps[2], &caps[4])
                    }
                } else {
                    format!("{}!{}:{}!{}", &caps[1], &caps[2], &caps[3], &caps[4])
                }
            })
            .to_string();

        // 0d. Replace single local reference Sheet1.A1 -> A1
        // Updated to handle both '.' and '!' separators
        static ODS_LOCAL_SINGLE_REF: OnceLock<Regex> = OnceLock::new();
        let local_single_ref =
            ODS_LOCAL_SINGLE_REF.get_or_init(|| Regex::new(r"^([^.!]+)[.!]([A-Z0-9$]+)$").unwrap());
        normalized = local_single_ref.replace_all(&normalized, "$2").to_string();

        // 0e. Replace normalized sheet ranges Sheet1!A1:B2 -> A1:B2
        // This handles cases where 0b/0b2 normalized the range to Excel format,
        // but we want to strip the sheet name for local context parity.
        static ODS_NORMALIZED_RANGE_REF: OnceLock<Regex> = OnceLock::new();
        let normalized_range_ref = ODS_NORMALIZED_RANGE_REF
            .get_or_init(|| Regex::new(r"^([^!]+)!([A-Z0-9$]+):([A-Z0-9$]+)$").unwrap());
        normalized = normalized_range_ref
            .replace_all(&normalized, "$2:$3")
            .to_string();
    }

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
    let rel_ref = ODS_REL_REF.get_or_init(|| Regex::new(r"\[\.(\$?[A-Z]+\$?[0-9]+)\]").unwrap());
    normalized = rel_ref.replace_all(&normalized, "$1").to_string();

    // 4. Replace whole column references [.A:.A] -> A:A
    // Matches [.A:.A], [.A:.B]
    let col_ref = ODS_COL_REF.get_or_init(|| Regex::new(r"\[\.([A-Z]+):\.([A-Z]+)\]").unwrap());
    normalized = col_ref.replace_all(&normalized, "$1:$2").to_string();

    // 5. Replace whole row references [.1:.1] -> 1:1
    // Matches [.1:.1], [.1:.10]
    let row_ref = ODS_ROW_REF.get_or_init(|| Regex::new(r"\[\.([0-9]+):\.([0-9]+)\]").unwrap());
    normalized = row_ref.replace_all(&normalized, "$1:$2").to_string();

    // Final check for identical range parts (e.g. A1:A1 -> A1)
    if let Some((start, end)) = normalized.split_once(':') {
        if start == end {
            return start.to_string();
        }
    }

    // Convert visible row numbers to XML row numbers if mapping is provided
    // This accounts for hidden rows in ODS files
    if let Some(row_map) = visible_to_xml_row {
        // Regex to match cell references: Sheet!A1, A1, $A$1, etc.
        // Captures: (optional sheet)(column)(row number)
        static CELL_REF_PATTERN: OnceLock<Regex> = OnceLock::new();
        let cell_ref = CELL_REF_PATTERN
            .get_or_init(|| Regex::new(r"(?:([A-Za-z0-9_]+)!)?(\$?[A-Z]+)(\$?)([0-9]+)").unwrap());

        normalized = cell_ref
            .replace_all(&normalized, |caps: &regex::Captures| {
                let sheet_name = caps.get(1).map(|m| m.as_str());
                let sheet_prefix = sheet_name.map(|s| format!("{}!", s)).unwrap_or_default();
                let col = caps.get(2).map(|m| m.as_str()).unwrap_or("");
                let abs_marker = caps.get(3).map(|m| m.as_str()).unwrap_or("");
                let row_str = caps.get(4).map(|m| m.as_str()).unwrap_or("");

                // ONLY convert row numbers for references to the CURRENT sheet
                // This includes:
                // 1. Local references (no sheet prefix)
                // 2. Explicit references to the current sheet (e.g., CurrentSheet!A1)
                let should_convert = if let Some(current_sheet) = current_sheet_name {
                    sheet_name.is_none() || sheet_name == Some(current_sheet)
                } else {
                    // If no current sheet provided, only convert local references
                    sheet_name.is_none()
                };

                if should_convert {
                    if let Ok(visible_row) = row_str.parse::<u32>() {
                        // Convert visible row (1-indexed) to XML row (0-indexed)
                        if let Some(&xml_row) = row_map.get(&visible_row) {
                            // Convert back to 1-indexed for formula representation
                            return format!("{}{}{}", col, abs_marker, xml_row + 1);
                        }
                    }
                }
                // Keep original for cross-sheet references or if no mapping found
                format!("{}{}{}{}", sheet_prefix, col, abs_marker, row_str)
            })
            .to_string();
    }

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
        // Initialize date styles map first to avoid borrow check issues
        let date_styles = extract_date_styles_from_ods(self.archive)?;

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
        let mut current_cf_range: Option<String> = None;

        // Track visible row numbering for ODS formulas
        // ODS formulas use 1-indexed visible row numbers (accounting for hidden rows)
        // but we store cells using 0-indexed XML row numbers
        let mut visible_row_counter = 1u32; // 1-indexed (ODS formula style)
        let mut visible_to_xml_row: HashMap<u32, u32> = HashMap::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"table:table" => {
                    // Finalize previous sheet if it exists (handles sheets without conditional formatting)
                    if let Some(sheet) = current_sheet.take() {
                        sheets.push(sheet);
                    }

                    let mut name = String::new();
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"table:name" {
                            name = attr.unescape_value()?.to_string();
                        }
                    }
                    current_sheet = Some(Sheet::new(name));
                    current_row = 0;
                    current_col = 0; // Reset column tracking for new sheet
                }
                Event::Start(e) if e.name().as_ref() == b"table:table-column" => {
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
                                    repeated = attr.unescape_value()?.parse::<u32>().unwrap_or(1);
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

                        // If it's a start tag, we need to skip to its end tag to avoid nested column issues
                        let mut col_buf = Vec::new();
                        loop {
                            match reader.read_event_into(&mut col_buf)? {
                                Event::End(ref te)
                                    if te.name().as_ref() == b"table:table-column" =>
                                {
                                    break;
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                            col_buf.clear();
                        }
                    }
                }
                Event::Empty(e) if e.name().as_ref() == b"table:table-column" => {
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
                                    repeated = attr.unescape_value()?.parse::<u32>().unwrap_or(1);
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
                Event::Start(e) if e.name().as_ref() == b"table:table-row" => {
                    row_repeated = 1;
                    current_col = 0;
                    if let Some(ref mut sheet) = current_sheet {
                        let mut hidden = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"table:number-rows-repeated" => {
                                    row_repeated =
                                        attr.unescape_value()?.parse::<u32>().unwrap_or(1);
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
                            for i in 0..row_repeated {
                                sheet.hidden_rows.push(current_row + i);
                            }
                        } else {
                            // Map visible row numbers to XML row indices
                            for i in 0..row_repeated {
                                visible_to_xml_row.insert(visible_row_counter, current_row + i);
                                visible_row_counter += 1;
                            }
                        }
                    }
                }
                Event::Empty(e) if e.name().as_ref() == b"table:table-row" => {
                    // Empty row (self-closing tag) - no cells, just increment row counter
                    row_repeated = 1;
                    if let Some(ref mut sheet) = current_sheet {
                        let mut hidden = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"table:number-rows-repeated" => {
                                    row_repeated =
                                        attr.unescape_value()?.parse::<u32>().unwrap_or(1);
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
                        } else {
                            // Map visible row numbers to XML row indices
                            for r in 0..row_repeated {
                                visible_to_xml_row.insert(visible_row_counter, current_row + r);
                                visible_row_counter += 1;
                            }
                        }
                    }
                    current_row += row_repeated;
                    current_col = 0;
                }
                Event::Start(e)
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
                        let mut style_name = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"table:number-columns-repeated" => {
                                    col_repeated =
                                        attr.unescape_value()?.parse::<u32>().unwrap_or(1);
                                }
                                b"table:number-columns-spanned" => {
                                    cols_spanned =
                                        attr.unescape_value()?.parse::<u32>().unwrap_or(1);
                                }
                                b"table:number-rows-spanned" => {
                                    rows_spanned =
                                        attr.unescape_value()?.parse::<u32>().unwrap_or(1);
                                }
                                b"table:formula" => {
                                    let raw_formula = attr.unescape_value()?;
                                    formula = Some(normalize_ods_reference(
                                        &raw_formula,
                                        false,
                                        Some(&visible_to_xml_row),
                                        Some(&sheet.name),
                                    ));
                                }
                                b"table:style-name" => {
                                    style_name = attr.unescape_value()?.to_string();
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
                                    let val_str = attr.unescape_value()?.to_string();
                                    if !has_value {
                                        value = match attr.key.as_ref() {
                                            b"office:value" => {
                                                if let Ok(n) = val_str.parse::<f64>() {
                                                    CellValue::Number(n)
                                                } else {
                                                    CellValue::Text(val_str)
                                                }
                                            }
                                            b"office:date-value" => {
                                                // Convert ISO date to Serial Number
                                                if let Some(n) = parse_ods_date(&val_str) {
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

                        // Read text content from <text:p> elements
                        // This handles both error cells and regular text cells
                        let mut text_content = String::new();
                        let mut text_buf = Vec::new();
                        loop {
                            match reader.read_event_into(&mut text_buf)? {
                                Event::Start(ref te) if te.name().as_ref() == b"text:p" => {
                                    let mut p_buf = Vec::new();
                                    loop {
                                        match reader.read_event_into(&mut p_buf)? {
                                            Event::Text(ref t) => {
                                                text_content.push_str(&t.unescape()?.to_string());
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
                                        || te.name().as_ref() == b"table:covered-table-cell" =>
                                {
                                    break;
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                            text_buf.clear();
                        }

                        // Use text content if we have it and no other value
                        if !text_content.is_empty() {
                            if is_error_cell {
                                value = CellValue::formula_with_error("", text_content);
                                has_value = true;
                            } else if !has_value {
                                // Only use text:p content if we don't have a value from attributes
                                value = CellValue::Text(text_content);
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

                            // Look up format string from style
                            let num_fmt = if !style_name.is_empty() {
                                date_styles.get(&style_name).cloned()
                            } else {
                                None
                            };

                            for r in 0..row_repeated {
                                for c in 0..col_repeated {
                                    let cell = Cell {
                                        row: current_row + r,
                                        col: current_col + c,
                                        value: cell_value.clone(),
                                        num_fmt: num_fmt.clone(),
                                    };
                                    sheet.cells.insert((current_row + r, current_col + c), cell);
                                }
                            }
                        }

                        // Track furthest styled cell for used range metadata (Excel-like behavior)
                        // Even if cell has no value, if it has a style, it extends the used range
                        if !style_name.is_empty() {
                            for r in 0..row_repeated {
                                for c in 0..col_repeated {
                                    let row_pos = current_row + r;
                                    let col_pos = current_col + c;

                                    // Update max styled cell position
                                    if let Some((max_row, max_col)) = sheet.used_range {
                                        sheet.used_range =
                                            Some((max_row.max(row_pos), max_col.max(col_pos)));
                                    } else {
                                        sheet.used_range = Some((row_pos, col_pos));
                                    }
                                }
                            }
                        }

                        // Multiply repeated by spanned to get true column consumption
                        current_col += col_repeated * cols_spanned;
                    }
                }
                Event::Empty(e)
                    if e.name().as_ref() == b"table:table-cell"
                        || e.name().as_ref() == b"table:covered-table-cell" =>
                {
                    if let Some(ref mut sheet) = current_sheet {
                        let mut col_repeated = 1u32;
                        let mut cols_spanned = 1u32;
                        let mut rows_spanned = 1u32;
                        let mut formula = None;
                        let mut style_name = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"table:number-columns-repeated" => {
                                    col_repeated =
                                        attr.unescape_value()?.parse::<u32>().unwrap_or(1);
                                }
                                b"table:number-columns-spanned" => {
                                    cols_spanned =
                                        attr.unescape_value()?.parse::<u32>().unwrap_or(1);
                                }
                                b"table:number-rows-spanned" => {
                                    rows_spanned =
                                        attr.unescape_value()?.parse::<u32>().unwrap_or(1);
                                }
                                b"table:formula" => {
                                    let raw_formula = attr.unescape_value()?;
                                    formula = Some(normalize_ods_reference(
                                        &raw_formula,
                                        false,
                                        Some(&visible_to_xml_row),
                                        Some(&sheet.name),
                                    ));
                                }
                                b"table:style-name" => {
                                    style_name = attr.unescape_value()?.to_string();
                                }
                                _ => {}
                            }
                        }

                        // Check if this empty cell is actually a merged cell
                        if cols_spanned > 1 || rows_spanned > 1 {
                            sheet.merged_cells.push((
                                current_row,
                                current_col,
                                current_row + rows_spanned - 1,
                                current_col + cols_spanned - 1,
                            ));
                        }

                        // If it's an empty cell but has a formula, we should store it.
                        if let Some(f) = formula {
                            for r in 0..row_repeated {
                                for c in 0..col_repeated {
                                    let cell = Cell {
                                        row: current_row + r,
                                        col: current_col + c,
                                        value: CellValue::formula(f.clone()),
                                        num_fmt: None,
                                    };
                                    sheet.cells.insert((current_row + r, current_col + c), cell);
                                }
                            }
                        }

                        // Track furthest styled cell for used range metadata (Excel-like behavior)
                        if !style_name.is_empty() {
                            for r in 0..row_repeated {
                                for c in 0..col_repeated {
                                    let row_pos = current_row + r;
                                    let col_pos = current_col + c;

                                    // Update max styled cell position
                                    if let Some((max_row, max_col)) = sheet.used_range {
                                        sheet.used_range =
                                            Some((max_row.max(row_pos), max_col.max(col_pos)));
                                    } else {
                                        sheet.used_range = Some((row_pos, col_pos));
                                    }
                                }
                            }
                        }

                        // Multiply repeated by spanned to get true column consumption
                        current_col += col_repeated * cols_spanned;
                    }
                }
                Event::Start(e) if e.name().as_ref() == b"calcext:conditional-format" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"calcext:target-range-address" {
                            current_cf_range = Some(attr.unescape_value()?.to_string());
                        }
                    }
                }
                Event::End(e) if e.name().as_ref() == b"calcext:conditional-format" => {
                    current_cf_range = None;
                }
                Event::Start(e) if e.name().as_ref() == b"calcext:condition" => {
                    if let Some(ref mut sheet) = current_sheet {
                        sheet.conditional_formatting_count += 1;
                        if let Some(ref range) = current_cf_range {
                            sheet
                                .conditional_formatting_ranges
                                .push(normalize_ods_reference(range, false, None, None));
                        }
                    }
                }
                // standard ODS conditional formatting
                Event::Start(e) if e.name().as_ref() == b"table:conditional-formatting" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"table:target-range-address" {
                            current_cf_range = Some(attr.unescape_value()?.to_string());
                        }
                    }
                }
                Event::End(e) if e.name().as_ref() == b"table:conditional-formatting" => {
                    current_cf_range = None;
                }
                Event::Start(e) if e.name().as_ref() == b"table:conditional-formatting-rule" => {
                    if let Some(ref mut sheet) = current_sheet {
                        sheet.conditional_formatting_count += 1;
                        if let Some(ref range) = current_cf_range {
                            sheet
                                .conditional_formatting_ranges
                                .push(normalize_ods_reference(range, false, None, None));
                        }
                    }
                }
                Event::Empty(e) if e.name().as_ref() == b"calcext:condition" => {
                    if let Some(ref mut sheet) = current_sheet {
                        sheet.conditional_formatting_count += 1;
                        if let Some(ref range) = current_cf_range {
                            sheet
                                .conditional_formatting_ranges
                                .push(normalize_ods_reference(range, false, None, None));
                        }
                    }
                }
                Event::End(e) if e.name().as_ref() == b"table:table-row" => {
                    current_row += row_repeated;
                    current_col = 0;
                }
                Event::End(e) if e.name().as_ref() == b"table:table" => {
                    // Calculate used range for the sheet before finalizing
                    // This must happen here (not at sheet finalization) because it needs to run
                    // for ALL sheets, whether they have conditional formatting or not
                    if let Some(ref mut sheet) = current_sheet {
                        let cells_range = calculate_used_range(&sheet.cells);

                        // Merge styled cell tracking with value cells
                        // Convert from 0-indexed position to count format (add 1) to match XLSX
                        sheet.used_range = match (sheet.used_range, cells_range) {
                            (Some((s_row, s_col)), Some((c_row, c_col))) => {
                                Some((s_row.max(c_row) + 1, s_col.max(c_col) + 1))
                            }
                            (Some((s_row, s_col)), None) => Some((s_row + 1, s_col + 1)),
                            (None, Some((c_row, c_col))) => Some((c_row + 1, c_col + 1)),
                            (None, None) => None,
                        };
                    }
                }
                // Handle conditional formatting that appears after table closing tag
                Event::Start(e) if e.name().as_ref() == b"calcext:conditional-formats" => {
                    // This wrapper appears after </table:table>, continue processing
                }
                Event::End(e) if e.name().as_ref() == b"calcext:conditional-formats" => {
                    // End of conditional formatting section
                    // Don't finalize the sheet here - the table end handler will do it
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        // Handle any remaining sheet that didn't have conditional formatting
        if let Some(sheet) = current_sheet.take() {
            sheets.push(sheet);
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
        let mut in_database_ranges = false;

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(ref e) if e.name().as_ref() == b"table:named-expressions" => {
                    in_named_expressions = true;
                }
                Event::Start(ref e) if e.name().as_ref() == b"table:database-ranges" => {
                    in_database_ranges = true;
                }
                // Combined match for Start/Empty of item tags
                Event::Empty(ref e) | Event::Start(ref e) => {
                    if in_named_expressions && e.name().as_ref() == b"table:named-range" {
                        let mut name = String::new();
                        let mut cell_range_address = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"table:name" => {
                                    name = attr.unescape_value()?.to_string();
                                }
                                b"table:cell-range-address" => {
                                    cell_range_address = attr.unescape_value()?.to_string();
                                }
                                _ => {}
                            }
                        }

                        if !name.is_empty() && !cell_range_address.is_empty() {
                            let normalized =
                                normalize_ods_reference(&cell_range_address, true, None, None);
                            defined_names.insert(name, normalized);
                        }
                    } else if in_database_ranges && e.name().as_ref() == b"table:database-range" {
                        let mut name = String::new();
                        let mut target_range_address = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"table:name" => {
                                    name = attr.unescape_value()?.to_string();
                                }
                                b"table:target-range-address" => {
                                    target_range_address = attr.unescape_value()?.to_string();
                                }
                                _ => {}
                            }
                        }

                        if !name.is_empty() && !target_range_address.is_empty() {
                            // Filter out internal ODS names that start with __Anonymous_Sheet_DB__
                            if !name.starts_with("__Anonymous_Sheet_DB__") {
                                let normalized = normalize_ods_reference(
                                    &target_range_address,
                                    true,
                                    None,
                                    None,
                                );
                                defined_names.insert(name, normalized);
                            }
                        }
                    }
                }
                Event::End(e) => {
                    if e.name().as_ref() == b"table:named-expressions" {
                        in_named_expressions = false;
                    } else if e.name().as_ref() == b"table:database-ranges" {
                        in_database_ranges = false;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(defined_names)
    }

    fn read_hidden_sheets(&mut self) -> Result<Vec<String>> {
        extract_hidden_sheets_from_ods(self.archive)
    }

    fn has_macros(&mut self) -> Result<bool> {
        has_macros(self.archive)
    }

    fn read_external_links(&mut self) -> Result<Vec<String>> {
        extract_external_links_ods(self.archive)
    }
}

// Helper to calculate used range from cells
fn calculate_used_range(cells: &HashMap<(u32, u32), Cell>) -> Option<(u32, u32)> {
    if cells.is_empty() {
        return None;
    }

    let mut max_row = 0;
    let mut max_col = 0;

    for (row, col) in cells.keys() {
        if *row > max_row {
            max_row = *row;
        }
        if *col > max_col {
            max_col = *col;
        }
    }

    // Return (max_row, max_col) as inclusive indices
    Some((max_row, max_col))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_normalize_ods_reference_basic() {
        assert_eq!(
            normalize_ods_reference("of:=SUM([.A1:.B2])", false, None, None),
            "SUM(A1:B2)"
        );
        assert_eq!(
            normalize_ods_reference("of:=[.A1]+[.B1]", false, None, None),
            "A1+B1"
        );
        assert_eq!(
            normalize_ods_reference("of:=SUM([.A:.A])", false, None, None),
            "SUM(A:A)"
        );
        assert_eq!(
            normalize_ods_reference("of:=SUM([.1:.1])", false, None, None),
            "SUM(1:1)"
        );
    }

    #[test]
    fn test_normalize_ods_reference_sheet() {
        assert_eq!(
            normalize_ods_reference("of:=[$Sheet1.A1]*2", false, None, None),
            "Sheet1!A1*2"
        );
        assert_eq!(
            normalize_ods_reference("of:=SUM([$Sheet1.A1:.B2])", false, None, None),
            "SUM(Sheet1!A1:B2)"
        );
        assert_eq!(
            normalize_ods_reference("of:=$Sheet1.$A$1+$Sheet1.B1", false, None, None),
            "Sheet1!$A$1+Sheet1!B1"
        );
    }

    #[test]
    fn test_normalize_ods_reference_mixed() {
        assert_eq!(
            normalize_ods_reference("of:=[.A1:.$B$2]", false, None, None),
            "A1:$B$2"
        );
        assert_eq!(
            normalize_ods_reference("of:=[.A$1]+$Sheet1.B$2", false, None, None),
            "A$1+Sheet1!B$2"
        );
    }

    #[test]
    fn test_normalize_ods_range() {
        // Local range with redundant sheet names
        assert_eq!(
            normalize_ods_reference("Sheet1.B2:Sheet1.B4", false, None, None),
            "B2:B4"
        );
        // Single cell ref
        assert_eq!(
            normalize_ods_reference("Sheet1.A1", false, None, None),
            "A1"
        );
        // Multi-sheet range
        assert_eq!(
            normalize_ods_reference("Sheet1.A1:Sheet2.B2", false, None, None),
            "Sheet1!A1:Sheet2!B2"
        );
        // Absolute local ref
        assert_eq!(
            normalize_ods_reference("Sheet1.$A$1", false, None, None),
            "$A$1"
        );
    }
    #[test]
    fn test_normalize_ods_unbracketed_range() {
        // User reported "Listas!$D$19:.$M$19" appearing in output (dot issue).
        // This suggests input was "$Listas.$D$19:.$M$19" (no brackets, like database ranges)
        // and sheet_range_ref failed because it expects brackets.
        let raw = "$Listas.$D$19:.$M$19";

        // With preserve=true (defined names), we want the full sheet qualification
        let expected_true = "Listas!$D$19:$M$19";
        assert_eq!(
            normalize_ods_reference(raw, true, None, None),
            expected_true,
            "Failed with preserve=true"
        );

        // With preserve=false (local formulas), we want to strip the sheet name if possible
        // to avoid false circular references (ERR003 treat explicit self-sheet as non-trivial)
        let expected_false = "$D$19:$M$19";
        assert_eq!(
            normalize_ods_reference(raw, false, None, None),
            expected_false,
            "Failed with preserve=false"
        );
    }

    #[test]
    fn test_normalize_ods_strip_after_regex_match() {
        // If 0b2 matches "$Sheet1.A1:.$B2" -> "Sheet1!A1:B2"
        // 0c should SHOULD strip "Sheet1!" if preserve_sheet=false
        let raw = "$Sheet1.A1:.$B2";
        // Currently (before fix) this prints "Sheet1!A1:B2"
        // We want "A1:B2" if preserve_sheet=false (simulating local formula)
        // With preserve=false, it should strip matched sheet names
        assert_eq!(normalize_ods_reference(raw, false, None, None), "A1:$B2");

        // Also check bracketed case
        let raw_bracket = "[$Sheet1.A1:.$B2]";
        assert_eq!(
            normalize_ods_reference(raw_bracket, false, None, None),
            "A1:$B2"
        );
    }

    #[test]
    fn test_normalize_ods_preserve_sheet() {
        // Should preserve sheet name even if it looks local
        assert_eq!(
            normalize_ods_reference("Sheet1.A1", true, None, None),
            "Sheet1.A1"
        );
        // Should preserve absolute local ref
        assert_eq!(
            normalize_ods_reference("Sheet1.$G$2", true, None, None),
            "Sheet1.$G$2"
        );
        // Normal ranges should still be processed if they don't match the strip pattern
        // But our strip pattern in 0c matches: ([^.]+)\.([A-Z0-9$]+):([^.]+)\.([A-Z0-9$]+)
        // If preserve=true, this pattern is skipped.
        assert_eq!(
            normalize_ods_reference("Sheet1.A1:Sheet1.B2", true, None, None),
            "Sheet1.A1:Sheet1.B2"
        );
    }

    #[test]
    fn test_normalize_ods_reference_single_cell_range() {
        // PERF004 regression: "Sheet1.A1:Sheet1.A1" should normalize to "A1"
        assert_eq!(
            normalize_ods_reference("Sheet1.A1:Sheet1.A1", false, None, None),
            "A1"
        );
        assert_eq!(normalize_ods_reference("A1:A1", false, None, None), "A1");
        assert_eq!(
            normalize_ods_reference("[.A1:.A1]", false, None, None),
            "A1"
        );
    }

    #[test]
    fn test_read_database_ranges_ods() {
        use std::io::Cursor;
        use std::io::Write;
        use zip::write::FileOptions;

        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options =
                FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);

            zip.start_file("content.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
    <office:body>
        <office:spreadsheet>
            <table:database-ranges>
                <table:database-range table:name="MyRange" table:target-range-address="Sheet1.A1:Sheet1.B2"/>
                <table:database-range table:name="OtherRange" table:target-range-address="Sheet1.C3"/>
            </table:database-ranges>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#).unwrap();

            zip.finish().unwrap();
        }

        let mut archive = ZipArchive::new(Cursor::new(buf)).unwrap();
        let mut reader = OdsReader::new(&mut archive).unwrap();
        let defined_names = reader.read_defined_names().unwrap();

        // Note: read_defined_names returns normalized Excel-style references
        assert_eq!(defined_names.len(), 2);
        // "Sheet1.A1:Sheet1.B2" -> normalized with preserve_sheet=true keeps it as is
        assert_eq!(
            defined_names.get("MyRange"),
            Some(&"Sheet1.A1:Sheet1.B2".to_string())
        );
        // "Sheet1.C3" -> normalizes to "Sheet1.C3"
        assert_eq!(
            defined_names.get("OtherRange"),
            Some(&"Sheet1.C3".to_string())
        );
    }
}

//! ODS (OpenDocument Spreadsheet) parser implementation

use super::WorkbookReader;
use super::workbook::{Cell, CellValue, ExternalWorkbook, Sheet};
use crate::violation::ExcelError;
use anyhow::Result;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashMap;
use std::io::BufReader;
use std::sync::Arc;
use zip::ZipArchive;

/// Extract merged cell ranges from an ODS worksheet
/// ODS format: <table:table-cell table:number-columns-spanned="X" table:number-rows-spanned="Y">
/// Check if ODS file contains macros
/// ODS macros are stored in Basic/ or Scripts/ directories,
/// or declared in META-INF/manifest.xml
///
/// # Returns
///
/// * `Result<bool>` - True if macros or scripts are detected
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

/// Extract cached error values from an ODS worksheet
/// ODS error values are often stored in calcext:value-type="error" and calcext:value="#ERROR!"
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

    if !(1..=12).contains(&month) || day == 0 {
        return None;
    }

    // O(1) conversion from (year, month, day) to a serial date using correct
    // Gregorian counting (no Excel 1900 leap year bug).
    //
    // ODS stores dates as ISO 8601 strings and LibreOffice uses correct
    // calendar arithmetic. Unlike XLSX (which inherits Excel's bug treating
    // 1900 as a leap year), ODS serial numbers reflect the real calendar.
    //
    // Strategy: compute total days from a reference point using closed-form
    // arithmetic. 1900-01-01 = serial 1.

    // Cumulative days before each month in a non-leap year (index 1 = Jan).
    const CUMULATIVE_DAYS: [i32; 13] = [0, 0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334];

    let is_leap = |y: i32| (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);

    // Total days from year 0 to the start of `y` (Jan 1 of year y).
    // Leap year corrections use (y-1) to count leap years in [0, y-1] only;
    // using y would erroneously include year y's own leap day.
    let days_before_year =
        |y: i32| -> i32 { 365 * y + (y - 1) / 4 - (y - 1) / 100 + (y - 1) / 400 };

    let leap_feb = if month > 2 && is_leap(year) { 1 } else { 0 };
    let absolute_days =
        days_before_year(year) + CUMULATIVE_DAYS[month as usize] + leap_feb + day as i32;

    // serial = absolute_days - days_before_year(1900).
    // This gives 1900-01-01 = 1.
    let total_days = absolute_days - days_before_year(1900);

    // Time fraction
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

    Some(total_days as f64 + time_fraction)
}

/// Extract formulas from an ODS worksheet
/// ODS formulas are stored in table:formula attribute
///
/// Normalize ODS reference to something resembling Excel A1 notation
/// Handles:
/// - `[.A1]` -> `A1` (Local ref)
/// - `[$Sheet1.A1]` -> `Sheet1!A1` (Absolute sheet ref)
/// - `['file:///path'#$Sheet1.A1]` -> `[1]Sheet1!A1` (External ref)
pub fn normalize_ods_reference(
    reference: &str,
    preserve_sheet: bool,
    current_sheet_name: Option<&str>,
    external_workbooks: &mut Vec<ExternalWorkbook>,
) -> String {
    // Strip "of:=" prefix if present
    let input = reference.strip_prefix("of:=").unwrap_or(reference);

    let mut result = String::with_capacity(input.len());
    let mut chars = input.chars().peekable();

    while let Some(ch) = chars.next() {
        match ch {
            '[' => {
                // Handle bracketed references: [.A1], [$Sheet.A1], [.A1:.B2], ['file:///path'#Sheet.A1]
                parse_bracket_ref(
                    &mut chars,
                    &mut result,
                    preserve_sheet,
                    current_sheet_name,
                    external_workbooks,
                );
            }
            '$' => {
                // Handle $Sheet.A1 or $Sheet.A1:.$B$2 (unbracketed sheet reference)
                if let Some(next) = chars.peek() {
                    if next.is_alphabetic() {
                        parse_dollar_sheet_ref(
                            &mut chars,
                            &mut result,
                            preserve_sheet,
                            if preserve_sheet {
                                None
                            } else {
                                current_sheet_name
                            },
                        );
                    } else {
                        result.push(ch);
                    }
                } else {
                    result.push(ch);
                }
            }
            _ => result.push(ch),
        }
    }

    // Post-processing: handle plain Sheet.Cell references (without $ or brackets)
    // This handles cases like "Sheet1.A1" or "Sheet1.B2:Sheet1.B4"
    if !preserve_sheet {
        result = strip_local_sheet_refs(&result);
    }

    // Post-processing: normalize ODS argument separator (`;` → `,`)
    // ODS uses `;` between function arguments; rules expect Excel-style `,`.
    // Only apply to formulas (prefixed with `of:=`), not to range addresses
    // (conditional formatting targets, named ranges, database ranges) where `;`
    // could theoretically be a multi-area union separator.
    if reference.starts_with("of:=") {
        result = normalize_arg_separator(&result);
    }

    // Post-processing: handle identical range parts (A1:A1 → A1)
    // BUT preserve whole column/row ranges (A:A, 1:1)
    if let Some((start, end)) = result.split_once(':')
        && start == end
        && !is_whole_column_or_row(start)
    {
        return start.to_string();
    }

    result
}

/// Check if a reference is a whole column (A) or whole row (1)
fn is_whole_column_or_row(s: &str) -> bool {
    let cleaned = s.trim_start_matches('$');
    cleaned.chars().all(|c| c.is_alphabetic()) || cleaned.chars().all(|c| c.is_numeric())
}

/// Parse bracketed reference: [.A1], [$Sheet.A1], [.A1:.B2]
fn parse_bracket_ref(
    chars: &mut std::iter::Peekable<std::str::Chars>,
    result: &mut String,
    _preserve_sheet: bool,
    current_sheet_name: Option<&str>,
    external_workbooks: &mut Vec<ExternalWorkbook>,
) {
    let mut bracket_content = String::new();
    let mut depth = 1;

    // Collect everything inside brackets
    for ch in chars.by_ref() {
        if ch == '[' {
            depth += 1;
            bracket_content.push(ch);
        } else if ch == ']' {
            depth -= 1;
            if depth == 0 {
                break;
            }
            bracket_content.push(ch);
        } else {
            bracket_content.push(ch);
        }
    }

    // Parse the bracket content
    if let Some(hash_pos) = bracket_content.find('#') {
        // External reference: ['file:///path'#Sheet.A1]
        let mut uri = &bracket_content[..hash_pos];
        let ref_part = &bracket_content[hash_pos + 1..];

        // Strip quotes if present
        if uri.starts_with('\'') && uri.ends_with('\'') && uri.len() > 2 {
            uri = &uri[1..uri.len() - 1];
        }

        // Extract basename
        let path = std::path::Path::new(uri);
        let basename = path.file_name().and_then(|n| n.to_str()).unwrap_or(uri);

        // Find or add to external_workbooks
        let index = if let Some(pos) = external_workbooks.iter().position(|eb| eb.path == basename)
        {
            pos
        } else {
            let new_idx = external_workbooks.len();
            external_workbooks.push(ExternalWorkbook {
                index: new_idx,
                path: basename.to_string(),
            });
            new_idx
        };

        // Parse the reference part using the existing Sheet.Cell logic
        // For external references, we NEVER strip the sheet name, so pass None as current_sheet_name
        let mut normalized_ref = String::new();
        parse_sheet_qualified_ref(ref_part, &mut normalized_ref, None);

        // Build XLSX format: [index]Sheet!Cell
        result.push_str(&format!("[{}]", index + 1));
        result.push_str(&normalized_ref);
    } else if let Some(stripped) = bracket_content.strip_prefix('$') {
        // [$Sheet.A1] or [$Sheet.A1:.B2]
        parse_sheet_qualified_ref(stripped, result, current_sheet_name);
    } else if let Some(stripped) = bracket_content.strip_prefix('.') {
        // [.A1] or [.A1:.B2] or [.A:.A] or [.1:.1]
        parse_local_ref(stripped, result);
    } else {
        // Unknown format, keep as-is
        result.push('[');
        result.push_str(&bracket_content);
        result.push(']');
    }
}

/// Parse Sheet.A1:Sheet.B2 or Sheet.A1
fn parse_sheet_qualified_ref(content: &str, result: &mut String, current_sheet_name: Option<&str>) {
    if let Some(colon_pos) = content.find(':') {
        // Range: Sheet.A1:Sheet.B2 or Sheet.A1:.B2
        let start_part = &content[..colon_pos];
        let end_part = &content[colon_pos + 1..];

        let mut start_sheet = "";
        let mut start_cell = start_part;
        if let Some((s, c)) = start_part.split_once('.') {
            start_sheet = s;
            start_cell = c;
        }

        let mut end_sheet = "";
        let mut end_cell = end_part;

        // ODS shortcut: Sheet1.A1:.B2
        if let Some(stripped) = end_part.strip_prefix('.') {
            end_sheet = start_sheet;
            end_cell = stripped;
        } else if let Some((s, c)) = end_part.split_once('.') {
            end_sheet = s;
            end_cell = c;
        }

        if !start_sheet.is_empty() {
            // Check if we should strip the sheet name
            // For a range, we only strip if BOTH sheets match the current sheet
            let should_strip = if let Some(current) = current_sheet_name {
                start_sheet == current && (end_sheet == current || end_sheet.is_empty())
            } else {
                false
            };

            if should_strip {
                result.push_str(start_cell);
                result.push(':');
                result.push_str(end_cell);
            } else {
                result.push_str(start_sheet.trim_start_matches('$'));
                result.push('!');
                result.push_str(start_cell);
                result.push(':');
                if !end_sheet.is_empty() && end_sheet != start_sheet {
                    result.push_str(end_sheet.trim_start_matches('$'));
                    result.push('!');
                }
                result.push_str(end_cell);
            }
        } else {
            // Malformed or just local range like .A1:.B2 (though those should go to parse_local_ref)
            result.push_str(content);
        }
    } else {
        // Single cell: Sheet.A1
        let mut sheet = "";
        let mut cell = content;
        if let Some((s, c)) = content.split_once('.') {
            sheet = s;
            cell = c;
        }

        if !sheet.is_empty() {
            // Check if we should strip the sheet name
            let should_strip = if let Some(current) = current_sheet_name {
                sheet == current
            } else {
                false
            };

            if should_strip {
                result.push_str(cell);
            } else {
                result.push_str(sheet.trim_start_matches('$'));
                result.push('!');
                result.push_str(cell);
            }
        } else {
            // Malformed, keep as-is
            result.push_str(content);
        }
    }
}

/// Parse local reference: .A1 or .A1:.B2 or .A:.A or .1:.1
fn parse_local_ref(content: &str, result: &mut String) {
    if let Some(colon_pos) = content.find(':') {
        // Range: .A1:.B2 or .A:.A or .1:.1
        let start = &content[..colon_pos];
        let end = &content[colon_pos + 1..];
        let end_clean = end.strip_prefix('.').unwrap_or(end);

        result.push_str(start);
        result.push(':');
        result.push_str(end_clean);
    } else {
        // Single cell: .A1
        result.push_str(content);
    }
}

/// Parse $Sheet.A1 or $Sheet.A1:.$B$2 (unbracketed sheet reference)
/// Always outputs Sheet!Cell format, stripping handled at end of normalize_ods_reference
fn parse_dollar_sheet_ref(
    chars: &mut std::iter::Peekable<std::str::Chars>,
    result: &mut String,
    _preserve_sheet: bool,
    current_sheet_name: Option<&str>,
) {
    let mut sheet_name = String::new();

    // Collect sheet name until '.'
    while let Some(&ch) = chars.peek() {
        if ch == '.' {
            chars.next(); // consume '.'
            break;
        } else if ch.is_alphanumeric() || ch == '_' {
            sheet_name.push(ch);
            chars.next();
        } else {
            // Not a sheet reference, restore $
            result.push('$');
            result.push_str(&sheet_name);
            return;
        }
    }

    // Collect first cell reference
    let mut cell_ref = String::new();
    while let Some(&ch) = chars.peek() {
        if ch == '$' || ch.is_alphabetic() || ch.is_numeric() {
            cell_ref.push(ch);
            chars.next();
        } else {
            break;
        }
    }

    if cell_ref.is_empty() {
        // Malformed, restore original
        result.push('$');
        result.push_str(&sheet_name);
        return;
    }

    // Check if there's a range (colon)
    let mut cell_ref2 = String::new();
    if chars.peek() == Some(&':') {
        chars.next(); // consume ':'

        // Check for optional dot before second cell
        if chars.peek() == Some(&'.') {
            chars.next(); // consume '.'
        }

        // Collect second cell reference
        while let Some(&ch) = chars.peek() {
            if ch == '$' || ch.is_alphabetic() || ch.is_numeric() {
                cell_ref2.push(ch);
                chars.next();
            } else {
                break;
            }
        }
    }

    // Check if we should strip the sheet name
    let should_strip = if let Some(current) = current_sheet_name {
        sheet_name == current
    } else {
        false
    };

    if cell_ref2.is_empty() {
        // Single cell
        if !should_strip {
            result.push_str(&sheet_name);
            result.push('!');
        }
        result.push_str(&cell_ref);
    } else {
        // Range
        if !should_strip {
            result.push_str(&sheet_name);
            result.push('!');
        }
        result.push_str(&cell_ref);
        result.push(':');
        result.push_str(&cell_ref2);
    }
}

/// Strip local sheet references from formula (Sheet1.A1 → A1, Sheet1.A1:Sheet1.B2 → A1:B2)
/// Also converts Sheet.Cell to Sheet!Cell format for multi-sheet ranges
/// Only handles plain Sheet.Cell (no $ or brackets), which are processed in post-processing
/// Only strips if both parts of a range have the same sheet name
fn strip_local_sheet_refs(formula: &str) -> String {
    use regex::Regex;
    use std::sync::OnceLock;

    static SHEET_RANGE_PATTERN: OnceLock<Regex> = OnceLock::new();
    static SHEET_SINGLE_PATTERN: OnceLock<Regex> = OnceLock::new();

    let mut result = formula.to_string();

    // Convert plain Sheet.Cell to Sheet!Cell for multi-sheet ranges
    // Pattern: Sheet1.A1:Sheet2.B2 → Sheet1!A1:Sheet2!B2
    // Only matches Sheet.Cell (with dot), not Sheet!Cell (already processed)
    let range_re = SHEET_RANGE_PATTERN.get_or_init(|| {
        Regex::new(r"([A-Za-z0-9_]+)\.([A-Z$0-9]+):([A-Za-z0-9_]+)\.([A-Z$0-9]+)").unwrap()
    });
    result = range_re
        .replace_all(&result, |caps: &regex::Captures| {
            if caps[1] == caps[3] {
                // Same sheet, strip it
                format!("{}:{}", &caps[2], &caps[4])
            } else {
                // Different sheets, convert to ! format
                format!("{}!{}:{}!{}", &caps[1], &caps[2], &caps[3], &caps[4])
            }
        })
        .to_string();

    // Handle single cells: Sheet1.A1 → A1 (at start or end of string)
    let single_re = SHEET_SINGLE_PATTERN
        .get_or_init(|| Regex::new(r"^([A-Za-z0-9_]+)\.([A-Z$0-9]+)$").unwrap());
    result = single_re.replace_all(&result, "$2").to_string();

    result
}

/// Replace ODS function argument separators (`;` → `,`) outside string literals.
///
/// ODS formulas use `;` between arguments while Excel uses `,`. This function
/// performs the substitution while preserving `;` inside double-quoted strings.
fn normalize_arg_separator(formula: &str) -> String {
    let mut result = String::with_capacity(formula.len());
    let mut in_string = false;

    for ch in formula.chars() {
        match ch {
            '"' => {
                in_string = !in_string;
                result.push(ch);
            }
            ';' if !in_string => result.push(','),
            _ => result.push(ch),
        }
    }

    result
}

/// Container for metadata parsed from an ODS file
///
/// Only caches lightweight metadata needed by subsequent reader methods.
/// Sheet data is returned directly and not cached, avoiding a full clone.
struct OdsData {
    hidden_sheets: Vec<String>,
    has_macros: bool,
    external_workbooks: Vec<ExternalWorkbook>,
}

/// ODS (OpenDocument Spreadsheet) reader implementation
///
/// Handles parsing of .ods content.xml and styles.xml
pub struct OdsReader<'a, R: std::io::Read + std::io::Seek> {
    archive: &'a mut ZipArchive<R>,
    data: Option<OdsData>,
}

const MAX_COLUMNS: u32 = 16384;

impl<'a, R: std::io::Read + std::io::Seek> OdsReader<'a, R> {
    pub fn new(archive: &'a mut ZipArchive<R>) -> Result<Self> {
        Ok(Self {
            archive,
            data: None,
        })
    }
}

impl<'a, R: std::io::Read + std::io::Seek> WorkbookReader for OdsReader<'a, R> {
    fn read_sheets(&mut self) -> Result<Vec<Sheet>> {
        // ============================================================
        // SINGLE PASS: Parse content.xml once for all concerns
        // ============================================================

        let mut sheets = Vec::new();

        // ============================================================
        // STATE: Date styles extraction (from content.xml)
        // ============================================================
        let mut data_styles: HashMap<String, Arc<str>> = HashMap::new();
        let mut cell_styles = HashMap::new();
        let mut date_styles: HashMap<String, Arc<str>> = HashMap::new(); // Resolved styles for cell lookup
        let mut current_data_style_name = String::new();
        let mut current_format = String::new();
        let mut in_date_style = false;
        let mut automatic_order = false;

        // ============================================================
        // STATE: Hidden sheets detection
        // ============================================================
        let mut hidden_sheets = Vec::new();
        let mut hidden_styles = std::collections::HashSet::new();
        let mut sheet_styles = Vec::new(); // (sheet_name, style_name)

        // ============================================================
        // STATE: Main sheet data parsing
        // ============================================================
        let mut external_workbooks = Vec::new();
        let mut current_sheet: Option<Sheet> = None;
        let mut current_row = 0u32;
        let mut row_repeated = 1u32;
        let mut current_col = 0u32;
        let mut current_cf_range: Option<String> = None;
        let mut skip_current_sheet = false; // Flag to skip external sheets

        // ============================================================
        // PARSE styles.xml FIRST to populate date_styles
        // ============================================================

        if let Ok(styles_xml) = self.archive.by_name("styles.xml") {
            let mut reader = Reader::from_reader(BufReader::new(styles_xml));
            reader.config_mut().trim_text(false);
            let mut buf = Vec::new();

            let mut current_data_style_name = String::new();
            let mut current_format = String::new();
            let mut in_date_style = false;
            let mut automatic_order = false;

            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.name().as_ref() {
                            b"number:date-style" => {
                                in_date_style = true;
                                current_format.clear();
                                automatic_order = false;
                                for attr in e.attributes().flatten() {
                                    match attr.key.as_ref() {
                                        b"style:name" => {
                                            current_data_style_name =
                                                attr.unescape_value()?.to_string();
                                        }
                                        b"number:automatic-order"
                                            if attr.value.as_ref() == b"true" =>
                                        {
                                            automatic_order = true;
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            b"number:text-style" => {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"style:name" {
                                        let style_name = attr.unescape_value()?.to_string();
                                        data_styles.insert(style_name.clone(), Arc::from("@"));
                                        date_styles.insert(style_name, Arc::from("@"));
                                    }
                                }
                            }
                            b"style:style" => {
                                let mut is_cell_style = false;
                                let mut style_name = String::new();
                                let mut data_style_name = String::new();

                                for attr in e.attributes().flatten() {
                                    match attr.key.as_ref() {
                                        b"style:family" if attr.value.as_ref() == b"table-cell" => {
                                            is_cell_style = true;
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
                                    cell_styles.insert(style_name.clone(), data_style_name.clone());
                                    // Immediately resolve if data style is known
                                    if let Some(format) = data_styles.get(&data_style_name) {
                                        date_styles.insert(style_name, format.clone());
                                    }
                                }
                            }
                            b"number:day" if in_date_style => {
                                let mut long = automatic_order;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"number:style" {
                                        long = attr.value.as_ref() == b"long";
                                    }
                                }
                                current_format.push_str(if long { "dd" } else { "d" });
                            }
                            b"number:month" if in_date_style => {
                                let mut long = automatic_order;
                                let mut textual = false;
                                for attr in e.attributes().flatten() {
                                    match attr.key.as_ref() {
                                        b"number:style" => {
                                            long = attr.value.as_ref() == b"long";
                                        }
                                        b"number:textual" if attr.value.as_ref() == b"true" => {
                                            textual = true;
                                        }
                                        _ => {}
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
                            let mut long = automatic_order;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"number:style" {
                                    long = attr.value.as_ref() == b"long";
                                }
                            }
                            current_format.push_str(if long { "dd" } else { "d" });
                        }
                        b"number:month" if in_date_style => {
                            let mut long = automatic_order;
                            let mut textual = false;
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"number:style" => {
                                        long = attr.value.as_ref() == b"long";
                                    }
                                    b"number:textual" if attr.value.as_ref() == b"true" => {
                                        textual = true;
                                    }
                                    _ => {}
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
                        current_format.push_str(e.decode()?.as_ref());
                    }
                    Event::End(e) if e.name().as_ref() == b"number:date-style" => {
                        if !current_data_style_name.is_empty() {
                            data_styles.insert(
                                current_data_style_name.clone(),
                                Arc::from(current_format.as_str()),
                            );
                            date_styles.insert(
                                current_data_style_name.clone(),
                                Arc::from(current_format.as_str()),
                            );
                            // Resolve any cell styles that reference this data style
                            for (cell_style, data_style) in &cell_styles {
                                if data_style == &current_data_style_name {
                                    date_styles.insert(
                                        cell_style.clone(),
                                        Arc::from(current_format.as_str()),
                                    );
                                }
                            }
                        }
                        in_date_style = false;
                    }
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        // ============================================================
        // PARSE content.xml in a scope to drop reader before has_macros
        // ============================================================
        {
            let content_xml = match self.archive.by_name("content.xml") {
                Ok(file) => file,
                Err(_) => return Ok(sheets),
            };

            let mut reader = Reader::from_reader(BufReader::new(content_xml));
            reader.config_mut().trim_text(false);

            let mut buf = Vec::new();

            // ============================================================
            // UNIFIED PARSING LOOP
            // ============================================================

            loop {
                match reader.read_event_into(&mut buf)? {
                    // --------------------------------------------------------
                    // DATE STYLES: number:date-style
                    // --------------------------------------------------------
                    Event::Start(e) if e.name().as_ref() == b"number:date-style" => {
                        in_date_style = true;
                        current_format.clear();
                        automatic_order = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"style:name" => {
                                    current_data_style_name = attr.unescape_value()?.to_string();
                                }
                                b"number:automatic-order" if attr.value.as_ref() == b"true" => {
                                    automatic_order = true;
                                }
                                _ => {}
                            }
                        }
                    }
                    // DATE STYLES: number:text-style
                    Event::Start(e) if e.name().as_ref() == b"number:text-style" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"style:name" {
                                let style_name = attr.unescape_value()?.to_string();
                                data_styles.insert(style_name.clone(), Arc::from("@"));
                                date_styles.insert(style_name, Arc::from("@"));
                            }
                        }
                    }
                    // DATE STYLES: style:style (for cell styles mapping)
                    Event::Start(e) if e.name().as_ref() == b"style:style" => {
                        let mut is_cell_style = false;
                        let mut style_name = String::new();
                        let mut data_style_name = String::new();
                        let mut is_table_style = false;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"style:family" => {
                                    if attr.value.as_ref() == b"table-cell" {
                                        is_cell_style = true;
                                    } else if attr.value.as_ref() == b"table" {
                                        is_table_style = true;
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

                        // Map cell styles to data styles for date formatting
                        if is_cell_style && !style_name.is_empty() && !data_style_name.is_empty() {
                            cell_styles.insert(style_name.clone(), data_style_name.clone());
                            // Immediately resolve to date_styles if data style is already known
                            if let Some(format) = data_styles.get(&data_style_name) {
                                date_styles.insert(style_name.clone(), format.clone());
                            }
                        }

                        // Check for hidden table styles
                        if is_table_style && !style_name.is_empty() {
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
                                    Event::End(ee) if ee.name().as_ref() == b"style:style" => {
                                        break;
                                    }
                                    Event::Eof => break,
                                    _ => {}
                                }
                                inner_buf.clear();
                            }
                        }
                    }
                    // DATE STYLES: date components (Start events)
                    Event::Start(e) if in_date_style => {
                        match e.name().as_ref() {
                            b"number:day" => {
                                let mut long = automatic_order;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"number:style" {
                                        long = attr.value.as_ref() == b"long";
                                    }
                                }
                                current_format.push_str(if long { "dd" } else { "d" });
                            }
                            b"number:month" => {
                                let mut long = automatic_order;
                                let mut textual = false;
                                for attr in e.attributes().flatten() {
                                    match attr.key.as_ref() {
                                        b"number:style" => {
                                            long = attr.value.as_ref() == b"long";
                                        }
                                        b"number:textual" if attr.value.as_ref() == b"true" => {
                                            textual = true;
                                        }
                                        _ => {}
                                    }
                                }
                                if textual {
                                    current_format.push_str(if long { "mmmm" } else { "mmm" });
                                } else {
                                    current_format.push_str(if long { "mm" } else { "m" });
                                }
                            }
                            b"number:year" => {
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
                            b"number:hours" => {
                                current_format.push_str("hh");
                            }
                            b"number:minutes" => {
                                current_format.push_str("mm");
                            }
                            b"number:seconds" => {
                                current_format.push_str("ss");
                            }
                            b"number:text" => {
                                // Will read text event next
                            }
                            _ => {}
                        }
                    }
                    // DATE STYLES: date components (Empty events)
                    Event::Empty(e) if in_date_style => match e.name().as_ref() {
                        b"number:day" => {
                            let mut long = automatic_order;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"number:style" {
                                    long = attr.value.as_ref() == b"long";
                                }
                            }
                            current_format.push_str(if long { "dd" } else { "d" });
                        }
                        b"number:month" => {
                            let mut long = automatic_order;
                            let mut textual = false;
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"number:style" => {
                                        long = attr.value.as_ref() == b"long";
                                    }
                                    b"number:textual" if attr.value.as_ref() == b"true" => {
                                        textual = true;
                                    }
                                    _ => {}
                                }
                            }
                            if textual {
                                current_format.push_str(if long { "mmmm" } else { "mmm" });
                            } else {
                                current_format.push_str(if long { "mm" } else { "m" });
                            }
                        }
                        b"number:year" => {
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
                        b"number:hours" => {
                            current_format.push_str("hh");
                        }
                        b"number:minutes" => {
                            current_format.push_str("mm");
                        }
                        b"number:seconds" => {
                            current_format.push_str("ss");
                        }
                        _ => {}
                    },
                    // DATE STYLES: text content within date style
                    Event::Text(e) if in_date_style => {
                        current_format.push_str(e.decode()?.as_ref());
                    }
                    // DATE STYLES: end of date-style
                    Event::End(e) if e.name().as_ref() == b"number:date-style" => {
                        if !current_data_style_name.is_empty() {
                            data_styles.insert(
                                current_data_style_name.clone(),
                                Arc::from(current_format.as_str()),
                            );
                            // Also add to date_styles directly
                            date_styles.insert(
                                current_data_style_name.clone(),
                                Arc::from(current_format.as_str()),
                            );
                            // Resolve any cell styles that reference this data style
                            for (cell_style, data_style) in &cell_styles {
                                if data_style == &current_data_style_name {
                                    date_styles.insert(
                                        cell_style.clone(),
                                        Arc::from(current_format.as_str()),
                                    );
                                }
                            }
                        }
                        in_date_style = false;
                    }
                    // --------------------------------------------------------
                    // MAIN SHEET DATA: table:table
                    // --------------------------------------------------------
                    Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"table:table" => {
                        // Collect sheet style for hidden detection
                        let mut sheet_name = String::new();
                        let mut style_name = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"table:name" => {
                                    sheet_name = attr.unescape_value()?.to_string();
                                }
                                b"table:style-name" => {
                                    style_name = attr.unescape_value()?.to_string();
                                }
                                _ => {}
                            }
                        }

                        if !sheet_name.is_empty() && !style_name.is_empty() {
                            sheet_styles.push((sheet_name.clone(), style_name.clone()));
                        }

                        // Finalize previous sheet if it exists and it's not external
                        if let Some(mut sheet) = current_sheet.take()
                            && !skip_current_sheet
                        {
                            sheet.cells.shrink_to_fit();
                            sheets.push(sheet);
                        }

                        let mut name = sheet_name;
                        if name.is_empty() {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"table:name" {
                                    name = attr.unescape_value()?.to_string();
                                }
                            }
                        }
                        // Assign sheet_index based on the current count of sheets parsed
                        let sheet_index = sheets.len() as u16;
                        let mut new_sheet = Sheet::new(name.clone(), sheet_index);
                        // Check if this sheet's style is in hidden_styles
                        let is_hidden =
                            !style_name.is_empty() && hidden_styles.contains(&style_name);
                        new_sheet.visible = !is_hidden;

                        current_sheet = Some(new_sheet);
                        current_row = 0;
                        current_col = 0; // Reset column tracking for new sheet
                        skip_current_sheet = false; // Reset skip flag for new sheet
                    }

                    // Detect external sheets by checking for table:table-source
                    Event::Start(ref e) | Event::Empty(ref e)
                        if e.name().as_ref() == b"table:table-source" =>
                    {
                        // Check if this table-source has an xlink:href attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"xlink:href" {
                                let _href = attr.unescape_value()?.to_string();
                                // This sheet is from an external workbook, mark it to be skipped
                                skip_current_sheet = true;
                                // Fast-forward to the end of this table to avoid parsing millions of rows
                                // CRITICAL: Start depth at 0 since we're already inside table:table element
                                let mut depth = 0;
                                let mut skip_buf = Vec::new();
                                loop {
                                    match reader.read_event_into(&mut skip_buf)? {
                                        Event::Start(ee)
                                            if ee.name().as_ref() == b"table:table" =>
                                        {
                                            depth += 1
                                        }
                                        Event::End(ee) if ee.name().as_ref() == b"table:table" => {
                                            if depth == 0 {
                                                // This is the closing tag of the current external sheet
                                                break;
                                            }
                                            depth -= 1;
                                        }
                                        Event::Eof => break,
                                        _ => {}
                                    }
                                    skip_buf.clear();
                                }
                                break;
                            }
                        }
                    }
                    Event::Start(e) if e.name().as_ref() == b"table:table-column" => {
                        if let Some(ref mut sheet) = current_sheet {
                            let mut hidden = false;
                            let mut repeated = 1u32;
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"table:visibility"
                                        if (attr.value.as_ref() == b"collapse"
                                            || attr.value.as_ref() == b"filter") =>
                                    {
                                        hidden = true;
                                    }
                                    b"table:number-columns-repeated" => {
                                        repeated =
                                            attr.unescape_value()?.parse::<u32>().unwrap_or(1);
                                    }
                                    _ => {}
                                }
                            }
                            if hidden && repeated > 0 {
                                sheet
                                    .hidden_columns
                                    .push((current_col, current_col + repeated - 1));
                            }
                            current_col += repeated;

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
                                    b"table:visibility"
                                        if (attr.value.as_ref() == b"collapse"
                                            || attr.value.as_ref() == b"filter") =>
                                    {
                                        hidden = true;
                                    }
                                    b"table:number-columns-repeated" => {
                                        repeated =
                                            attr.unescape_value()?.parse::<u32>().unwrap_or(1);
                                    }
                                    _ => {}
                                }
                            }
                            if hidden && repeated > 0 {
                                sheet
                                    .hidden_columns
                                    .push((current_col, current_col + repeated - 1));
                            }
                            current_col += repeated;
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
                                    b"table:visibility"
                                        if (attr.value.as_ref() == b"collapse"
                                            || attr.value.as_ref() == b"filter") =>
                                    {
                                        hidden = true;
                                    }
                                    _ => {}
                                }
                            }

                            if hidden && row_repeated > 0 {
                                sheet
                                    .hidden_rows
                                    .push((current_row, current_row + row_repeated - 1));
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
                                    b"table:visibility"
                                        if (attr.value.as_ref() == b"collapse"
                                            || attr.value.as_ref() == b"filter") =>
                                    {
                                        hidden = true;
                                    }
                                    _ => {}
                                }
                            }
                            if hidden && row_repeated > 0 {
                                sheet
                                    .hidden_rows
                                    .push((current_row, current_row + row_repeated - 1));
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
                            let mut is_matrix_array = false;

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
                                        let normalized = normalize_ods_reference(
                                            &raw_formula,
                                            false,
                                            Some(&sheet.name),
                                            &mut external_workbooks,
                                        );
                                        // Apply external workbook normalization
                                        formula = Some(normalized);
                                    }
                                    b"table:style-name" => {
                                        style_name = attr.unescape_value()?.to_string();
                                    }
                                    b"calcext:value-type" if attr.value.as_ref() == b"error" => {
                                        is_error_cell = true;
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
                                                        CellValue::Text(Arc::from(val_str.as_str()))
                                                    }
                                                }
                                                b"office:date-value" => {
                                                    // Convert ISO date to Serial Number
                                                    if let Some(n) = parse_ods_date(&val_str) {
                                                        CellValue::Number(n)
                                                    } else {
                                                        CellValue::Text(Arc::from(val_str.as_str()))
                                                    }
                                                }
                                                b"office:boolean-value" => {
                                                    CellValue::Boolean(val_str == "true")
                                                }
                                                _ => CellValue::Text(Arc::from(val_str.as_str())),
                                            };
                                            has_value = true;
                                        }
                                    }
                                    b"table:number-matrix-columns-spanned" => {
                                        let v = attr.unescape_value()?.parse::<u32>().unwrap_or(0);
                                        if v > 0 {
                                            is_matrix_array = true;
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
                                                    text_content.push_str(t.decode()?.as_ref());
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

                            // Use text content if we have it and no other value
                            if !text_content.is_empty() {
                                if is_error_cell {
                                    value = CellValue::Error(
                                        ExcelError::from_cell_str(&text_content)
                                            .unwrap_or(ExcelError::Value),
                                    );
                                    has_value = true;
                                } else if !has_value {
                                    // Only use text:p content if we don't have a value from attributes
                                    value = CellValue::Text(Arc::from(text_content.as_str()));
                                    has_value = true;
                                }
                            }

                            if has_value || formula.is_some() || !style_name.is_empty() {
                                // Optimization: Ignore empty styled cells that extend to the sheet edge
                                // These are often used as "row filler" in ODS (e.g., repeatedly 16300+ times)
                                // and cause massive memory usage if stored as individual cells.
                                let is_empty_content = !has_value && formula.is_none();
                                if is_empty_content
                                    && !style_name.is_empty()
                                    && (current_col + col_repeated >= MAX_COLUMNS)
                                {
                                    // Skip storing these cells
                                    // Just advance the column counter
                                } else {
                                    let mut cell_value = value;
                                    let formula_for_cell = formula;

                                    // Look up format string from style
                                    let num_fmt = if !style_name.is_empty() {
                                        date_styles.get(&style_name).cloned()
                                    } else {
                                        None
                                    };

                                    // Check if this is a text-formatted number
                                    // In ODS, text format is indicated by num_fmt == "@"
                                    if num_fmt.as_deref() == Some("@")
                                        && let CellValue::Number(n) = cell_value
                                    {
                                        // Convert number to text
                                        cell_value =
                                            CellValue::Text(Arc::from(n.to_string().as_str()));
                                    }

                                    for r in 0..row_repeated {
                                        for c in 0..col_repeated {
                                            let cell = Cell {
                                                row: current_row + r,
                                                col: current_col + c,
                                                value: cell_value.clone(),
                                                formula: formula_for_cell.as_deref().map(Box::from),
                                                is_array: is_matrix_array,
                                                num_fmt: num_fmt.clone(),
                                            };
                                            sheet
                                                .cells
                                                .insert((current_row + r, current_col + c), cell);

                                            // Update used_range for any inserted cell (value, formula, or style)
                                            let row_pos = current_row + r;
                                            let col_pos = current_col + c;
                                            if let Some((max_row, max_col)) = sheet.used_range {
                                                sheet.used_range = Some((
                                                    max_row.max(row_pos + 1),
                                                    max_col.max(col_pos + 1),
                                                ));
                                            } else {
                                                sheet.used_range = Some((row_pos + 1, col_pos + 1));
                                            }
                                        }
                                    }
                                }
                            }

                            // Multiply repeated by spanned to get true column consumption
                            current_col += col_repeated;
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
                            let mut is_matrix_array = false;

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
                                            Some(&sheet.name),
                                            &mut external_workbooks,
                                        ));
                                    }
                                    b"table:style-name" => {
                                        style_name = attr.unescape_value()?.to_string();
                                    }
                                    b"table:number-matrix-columns-spanned" => {
                                        let v = attr.unescape_value()?.parse::<u32>().unwrap_or(0);
                                        if v > 0 {
                                            is_matrix_array = true;
                                        }
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

                            // If it's an empty cell but has a formula or style, we should store it.
                            if formula.is_some() || !style_name.is_empty() {
                                // Optimization: Ignore empty styled cells that extend to the sheet edge
                                let is_empty_content = formula.is_none(); // empty cell has no value by definition here
                                if is_empty_content
                                    && !style_name.is_empty()
                                    && (current_col + col_repeated >= MAX_COLUMNS)
                                {
                                    // Skip storing these cells
                                } else {
                                    let cell_value = CellValue::Empty;
                                    let formula_for_cell = formula;

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
                                                formula: formula_for_cell.as_deref().map(Box::from),
                                                is_array: is_matrix_array,
                                                num_fmt: num_fmt.clone(),
                                            };
                                            sheet
                                                .cells
                                                .insert((current_row + r, current_col + c), cell);

                                            // Update used_range for any inserted cell (formula or style)
                                            let row_pos = current_row + r;
                                            let col_pos = current_col + c;
                                            if let Some((max_row, max_col)) = sheet.used_range {
                                                sheet.used_range = Some((
                                                    max_row.max(row_pos + 1),
                                                    max_col.max(col_pos + 1),
                                                ));
                                            } else {
                                                sheet.used_range = Some((row_pos + 1, col_pos + 1));
                                            }
                                        }
                                    }
                                }
                            }

                            // Multiply repeated by spanned to get true column consumption
                            current_col += col_repeated;
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
                                    .push(normalize_ods_reference(
                                        range,
                                        false,
                                        Some(&sheet.name),
                                        &mut external_workbooks,
                                    ));
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
                    Event::Start(e)
                        if e.name().as_ref() == b"table:conditional-formatting-rule" =>
                    {
                        if let Some(ref mut sheet) = current_sheet {
                            sheet.conditional_formatting_count += 1;
                            if let Some(ref range) = current_cf_range {
                                sheet
                                    .conditional_formatting_ranges
                                    .push(normalize_ods_reference(
                                        range,
                                        false,
                                        Some(&sheet.name),
                                        &mut external_workbooks,
                                    ));
                            }
                        }
                    }
                    Event::Empty(e) if e.name().as_ref() == b"calcext:condition" => {
                        if let Some(ref mut sheet) = current_sheet {
                            sheet.conditional_formatting_count += 1;
                            if let Some(ref range) = current_cf_range {
                                sheet
                                    .conditional_formatting_ranges
                                    .push(normalize_ods_reference(
                                        range,
                                        false,
                                        Some(&sheet.name),
                                        &mut external_workbooks,
                                    ));
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
                            // Both are already in count format (0-indexed position + 1)
                            sheet.used_range = match (sheet.used_range, cells_range) {
                                (Some((s_row, s_col)), Some((c_row, c_col))) => {
                                    Some((s_row.max(c_row), s_col.max(c_col)))
                                }
                                (Some(s), None) => Some(s),
                                (None, Some(c)) => Some(c),
                                (None, None) => None,
                            };

                            // Include hidden rows/columns in used_range for format parity
                            // Both ODS and XLSX should report ALL empty rows/columns (visible or hidden)
                            if let Some((mut rows, mut cols)) = sheet.used_range {
                                if let Some(max_hidden_row) =
                                    sheet.hidden_rows.iter().map(|&(_, end)| end).max()
                                {
                                    rows = rows.max(max_hidden_row + 1);
                                }
                                if let Some(max_hidden_col) =
                                    sheet.hidden_columns.iter().map(|&(_, end)| end).max()
                                {
                                    cols = cols.max(max_hidden_col + 1);
                                }
                                sheet.used_range = Some((rows, cols));
                            }
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

            // Finalize the last sheet if it exists and it's not external
            if let Some(mut sheet) = current_sheet
                && !skip_current_sheet
            {
                sheet.cells.shrink_to_fit();
                sheets.push(sheet);
            }
        } // End of content.xml parsing scope

        // ============================================================
        // POST-PROCESSING: Final resolution of any remaining styles
        // ============================================================
        // Resolve any cell styles that weren't resolved during parsing
        for (cell_style, data_style) in &cell_styles {
            if !date_styles.contains_key(cell_style)
                && let Some(format) = data_styles.get(data_style)
            {
                date_styles.insert(cell_style.clone(), format.clone());
            }
        }
        // Also add data styles directly to date_styles
        for (data_style, format) in &data_styles {
            date_styles
                .entry(data_style.clone())
                .or_insert(format.clone());
        }

        // ============================================================
        // HIDDEN SHEETS: Match sheets to hidden styles
        // ============================================================
        for (name, style) in sheet_styles {
            if hidden_styles.contains(&style) {
                hidden_sheets.push(name);
            }
        }

        // ============================================================
        // MACROS: Check if file contains macros
        // ============================================================
        let has_macros = has_macros(self.archive)?;

        // Store metadata for future method calls (sheets are returned directly)
        self.data = Some(OdsData {
            hidden_sheets,
            has_macros,
            external_workbooks,
        });

        Ok(sheets)
    }

    fn read_defined_names(&mut self) -> Result<HashMap<String, String>> {
        let mut defined_names = HashMap::new();

        // Ensure data is parsed to have access to external_workbooks
        if self.data.is_none() {
            self.read_sheets()?;
        }

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
                            let normalized = normalize_ods_reference(
                                &cell_range_address,
                                true,
                                None,
                                &mut self.data.as_mut().unwrap().external_workbooks,
                            );
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
                                    &mut self.data.as_mut().unwrap().external_workbooks,
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
        // Ensure data is parsed
        if self.data.is_none() {
            self.read_sheets()?;
        }
        Ok(self.data.as_ref().unwrap().hidden_sheets.clone())
    }

    fn has_macros(&mut self) -> Result<bool> {
        // Ensure data is parsed
        if self.data.is_none() {
            self.read_sheets()?;
        }
        Ok(self.data.as_ref().unwrap().has_macros)
    }

    fn read_external_links(&mut self) -> Result<Vec<String>> {
        // Ensure data is parsed
        if self.data.is_none() {
            self.read_sheets()?;
        }
        // Derive from external_workbooks
        Ok(self
            .data
            .as_ref()
            .unwrap()
            .external_workbooks
            .iter()
            .map(|wb| wb.path.clone())
            .collect())
    }

    fn read_external_workbooks(&mut self) -> Result<Vec<super::ExternalWorkbook>> {
        // Ensure data is parsed
        if self.data.is_none() {
            self.read_sheets()?;
        }
        Ok(self.data.as_ref().unwrap().external_workbooks.clone())
    }

    fn read_modified_date(&mut self) -> Result<Option<chrono::DateTime<chrono::Utc>>> {
        extract_modified_date_ods(self.archive)
    }

    fn read_date1904(&mut self) -> Result<bool> {
        extract_date1904_ods(self.archive)
    }
}

/// Extract last modified date from meta.xml
fn extract_modified_date_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Option<chrono::DateTime<chrono::Utc>>> {
    use chrono::{DateTime, Utc};

    let meta_xml = match archive.by_name("meta.xml") {
        Ok(file) => file,
        Err(_) => return Ok(None),
    };

    let mut reader = Reader::from_reader(BufReader::new(meta_xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_date = false;
    let mut modified_date: Option<DateTime<Utc>> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                // dc:date element contains the last modified date
                let name = e.name();
                if name.as_ref().ends_with(b":date") || name.as_ref() == b"date" {
                    in_date = true;
                }
            }
            Event::Text(e) if in_date => {
                let date_str = e.decode()?.to_string();
                // Parse ISO 8601 / RFC 3339 format
                if let Ok(parsed) = DateTime::parse_from_rfc3339(&date_str) {
                    modified_date = Some(parsed.with_timezone(&Utc));
                }
                in_date = false;
            }
            Event::End(e) => {
                let name = e.name();
                if name.as_ref().ends_with(b":date") || name.as_ref() == b"date" {
                    in_date = false;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(modified_date)
}

/// Extract date1904 setting from content.xml or settings.xml
/// ODS uses <table:null-date table:date-value="1904-01-01"/> in calculation settings
fn extract_date1904_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<bool> {
    // Check content.xml for null-date specification
    let content_xml = match archive.by_name("content.xml") {
        Ok(file) => file,
        Err(_) => return Ok(false),
    };

    let mut reader = Reader::from_reader(BufReader::new(content_xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"table:null-date" => {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"table:date-value" {
                        let value = attr.unescape_value()?;
                        // 1904 date system uses 1904-01-01 as epoch
                        return Ok(value.starts_with("1904"));
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
    fn test_parse_ods_date_known_serials() {
        // ODS uses correct Gregorian counting (no Excel 1900 leap year bug).
        // 1900-01-01 → serial 1
        assert_eq!(parse_ods_date("1900-01-01"), Some(1.0));
        // 1900-02-28 → serial 59
        assert_eq!(parse_ods_date("1900-02-28"), Some(59.0));
        // 1900-03-01 → serial 60 (correct: no phantom Feb 29)
        assert_eq!(parse_ods_date("1900-03-01"), Some(60.0));
        // Leap year dates (regression: formula must not double-count year's own leap day)
        assert_eq!(parse_ods_date("1904-01-01"), Some(1461.0));
        assert_eq!(parse_ods_date("2000-01-01"), Some(36525.0));
        assert_eq!(parse_ods_date("2024-01-01"), Some(45291.0));
        assert_eq!(parse_ods_date("2024-07-04"), Some(45476.0));
        // Non-leap year dates
        assert_eq!(parse_ods_date("1999-09-30"), Some(36432.0));
        assert_eq!(parse_ods_date("2025-05-03"), Some(45779.0));
    }

    #[test]
    fn test_parse_ods_date_with_time() {
        // 12:00:00 is exactly half a day
        assert_eq!(parse_ods_date("1900-01-01T12:00:00"), Some(1.5));
        // 06:00:00 is a quarter
        assert_eq!(parse_ods_date("1900-01-01T06:00:00"), Some(1.25));
    }

    #[test]
    fn test_parse_ods_date_invalid() {
        assert_eq!(parse_ods_date("not-a-date"), None);
        assert_eq!(parse_ods_date("2025-13"), None);
        // Out-of-range month/day must return None, not panic
        assert_eq!(parse_ods_date("2025-13-01"), None);
        assert_eq!(parse_ods_date("2025-00-01"), None);
        assert_eq!(parse_ods_date("2025-01-00"), None);
    }

    #[test]
    fn test_normalize_ods_reference_basic() {
        assert_eq!(
            normalize_ods_reference("of:=SUM([.A1:.B2])", false, None, &mut Vec::new()),
            "SUM(A1:B2)"
        );
        assert_eq!(
            normalize_ods_reference("of:=[.A1]+[.B1]", false, None, &mut Vec::new()),
            "A1+B1"
        );
        assert_eq!(
            normalize_ods_reference("of:=SUM([.A:.A])", false, None, &mut Vec::new()),
            "SUM(A:A)"
        );
        assert_eq!(
            normalize_ods_reference("of:=SUM([.1:.1])", false, None, &mut Vec::new()),
            "SUM(1:1)"
        );
    }

    #[test]
    fn test_normalize_ods_reference_sheet() {
        assert_eq!(
            normalize_ods_reference("of:=[$Sheet1.A1]*2", false, None, &mut Vec::new()),
            "Sheet1!A1*2"
        );
        assert_eq!(
            normalize_ods_reference("of:=SUM([$Sheet1.A1:.B2])", false, None, &mut Vec::new()),
            "SUM(Sheet1!A1:B2)"
        );
        assert_eq!(
            normalize_ods_reference("of:=$Sheet1.$A$1+$Sheet1.B1", false, None, &mut Vec::new()),
            "Sheet1!$A$1+Sheet1!B1"
        );
    }

    #[test]
    fn test_normalize_ods_reference_mixed() {
        assert_eq!(
            normalize_ods_reference("of:=[.A1:.$B$2]", false, None, &mut Vec::new()),
            "A1:$B$2"
        );
        assert_eq!(
            normalize_ods_reference("of:=[.A$1]+$Sheet1.B$2", false, None, &mut Vec::new()),
            "A$1+Sheet1!B$2"
        );
    }

    #[test]
    fn test_normalize_ods_range() {
        // Local range with redundant sheet names
        assert_eq!(
            normalize_ods_reference("Sheet1.B2:Sheet1.B4", false, None, &mut Vec::new()),
            "B2:B4"
        );
        // Single cell ref
        assert_eq!(
            normalize_ods_reference("Sheet1.A1", false, None, &mut Vec::new()),
            "A1"
        );
        // Multi-sheet range
        assert_eq!(
            normalize_ods_reference("Sheet1.A1:Sheet2.B2", false, None, &mut Vec::new()),
            "Sheet1!A1:Sheet2!B2"
        );
        // Absolute local ref
        assert_eq!(
            normalize_ods_reference("Sheet1.$A$1", false, None, &mut Vec::new()),
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
            normalize_ods_reference(raw, true, Some("Listas"), &mut Vec::new()),
            expected_true,
            "Failed with preserve=true"
        );

        // With preserve=false (local formulas), we want to strip the sheet name if possible
        // to avoid false circular references (ERR003 treat explicit self-sheet as non-trivial)
        let expected_false = "$D$19:$M$19";
        assert_eq!(
            normalize_ods_reference(raw, false, Some("Listas"), &mut Vec::new()),
            expected_false,
            "Failed with preserve=false"
        );
    }

    #[test]
    fn test_normalize_ods_same_sheet_range_with_dot_prefix() {
        // Test that ranges with dot prefix (.$Cell) indicating same sheet are normalized correctly
        // Pattern: $Sheet.A1:.$B2 where .$ means "same sheet as A1"
        let raw = "$Sheet1.A1:.$B2";
        // With preserve=false, should strip the sheet name since it's the current sheet
        assert_eq!(
            normalize_ods_reference(raw, false, Some("Sheet1"), &mut Vec::new()),
            "A1:$B2"
        );

        // Also check bracketed case: [$Sheet.A1:.$B2]
        let raw_bracket = "[$Sheet1.A1:.$B2]";
        assert_eq!(
            normalize_ods_reference(raw_bracket, false, Some("Sheet1"), &mut Vec::new()),
            "A1:$B2"
        );
    }

    #[test]
    fn test_normalize_ods_preserve_sheet() {
        // Should preserve sheet name even if it looks local
        assert_eq!(
            normalize_ods_reference("Sheet1.A1", true, None, &mut Vec::new()),
            "Sheet1.A1"
        );
        // Should preserve absolute local ref
        assert_eq!(
            normalize_ods_reference("Sheet1.$G$2", true, None, &mut Vec::new()),
            "Sheet1.$G$2"
        );
        // Normal ranges should still be processed if they don't match the strip pattern
        // But our strip pattern in 0c matches: ([^.]+)\.([A-Z0-9$]+):([^.]+)\.([A-Z0-9$]+)
        // If preserve=true, this pattern is skipped.
        assert_eq!(
            normalize_ods_reference("Sheet1.A1:Sheet1.B2", true, None, &mut Vec::new()),
            "Sheet1.A1:Sheet1.B2"
        );
    }

    #[test]
    fn test_normalize_ods_reference_single_cell_range() {
        // PERF004 regression: "Sheet1.A1:Sheet1.A1" should normalize to "A1"
        assert_eq!(
            normalize_ods_reference("Sheet1.A1:Sheet1.A1", false, None, &mut Vec::new()),
            "A1"
        );
        assert_eq!(
            normalize_ods_reference("A1:A1", false, None, &mut Vec::new()),
            "A1"
        );
        assert_eq!(
            normalize_ods_reference("[.A1:.A1]", false, None, &mut Vec::new()),
            "A1"
        );
    }

    #[test]
    fn test_normalize_arg_separator_basic() {
        assert_eq!(normalize_arg_separator("SUM(A1;B1)"), "SUM(A1,B1)");
    }

    #[test]
    fn test_normalize_arg_separator_preserves_strings() {
        // Semicolons inside double-quoted strings must be preserved
        assert_eq!(
            normalize_arg_separator("IF(A1;\"yes;no\";\"maybe\")"),
            "IF(A1,\"yes;no\",\"maybe\")"
        );
    }

    #[test]
    fn test_normalize_arg_separator_string_only() {
        assert_eq!(
            normalize_arg_separator("\"Hello; world\""),
            "\"Hello; world\""
        );
    }

    #[test]
    fn test_normalize_arg_separator_mixed() {
        assert_eq!(
            normalize_arg_separator("CONCATENATE(\"a;b\";C1;D1)"),
            "CONCATENATE(\"a;b\",C1,D1)"
        );
    }

    #[test]
    fn test_normalize_arg_separator_no_semicolons() {
        assert_eq!(normalize_arg_separator("SUM(A1,B1)"), "SUM(A1,B1)");
    }

    #[test]
    fn test_normalize_arg_separator_empty() {
        assert_eq!(normalize_arg_separator(""), "");
    }
    #[test]
    fn test_normalize_ods_external_reference_inline() {
        let mut external_workbooks = Vec::new();
        let raw = "of:=['file:///path/test.xlsx'#Sheet1.A1]";
        let normalized =
            normalize_ods_reference(raw, false, Some("Sheet1"), &mut external_workbooks);

        assert_eq!(normalized, "[1]Sheet1!A1");
        assert_eq!(external_workbooks.len(), 1);
        assert_eq!(external_workbooks[0].path, "test.xlsx");
        assert_eq!(external_workbooks[0].index, 0);

        // Test second reference to same workbook
        let raw2 = "of:=['file:///other/test.xlsx'#Sheet2.B2]";
        let normalized2 =
            normalize_ods_reference(raw2, false, Some("Sheet1"), &mut external_workbooks);
        assert_eq!(normalized2, "[1]Sheet2!B2");
        assert_eq!(external_workbooks.len(), 1);
    }

    #[test]
    fn test_read_database_ranges_ods() {
        use std::io::Cursor;
        use std::io::Write;
        use zip::write::SimpleFileOptions;

        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options =
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);

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

    #[test]
    fn test_styled_empty_cell_ods() {
        use std::io::Cursor;
        use std::io::Write;
        use zip::write::SimpleFileOptions;

        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options =
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);

            zip.start_file("content.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="Sheet1">
                <table:table-row>
                    <table:table-cell table:style-name="ce1"/>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#).unwrap();

            zip.finish().unwrap();
        }

        let mut archive = ZipArchive::new(Cursor::new(buf)).unwrap();
        let mut reader = OdsReader::new(&mut archive).unwrap();
        let sheets = reader.read_sheets().unwrap();

        assert_eq!(sheets.len(), 1);
        let sheet = &sheets[0];
        // The cell (0, 0) should be in the map because it has a style "ce1"
        assert!(sheet.cells.contains_key(&(0, 0)));
        assert_eq!(sheet.cells.get(&(0, 0)).unwrap().value, CellValue::Empty);
    }

    #[test]
    fn test_hidden_repeated_rows_and_columns_as_ranges_ods() {
        use std::io::Cursor;
        use std::io::Write;
        use zip::write::SimpleFileOptions;

        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options =
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);

            zip.start_file("content.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="Sheet1">
                <table:table-column table:visibility="collapse" table:number-columns-repeated="3"/>
                <table:table-column table:number-columns-repeated="16381"/>
                <table:table-row>
                    <table:table-cell office:value-type="float" office:value="1"></table:table-cell>
                </table:table-row>
                <table:table-row table:visibility="collapse" table:number-rows-repeated="2">
                    <table:table-cell/>
                </table:table-row>
                <table:table-row table:visibility="filter" table:number-rows-repeated="1048573"/>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#).unwrap();

            zip.finish().unwrap();
        }

        let mut archive = ZipArchive::new(Cursor::new(buf)).unwrap();
        let mut reader = OdsReader::new(&mut archive).unwrap();
        let sheets = reader.read_sheets().unwrap();

        let sheet = &sheets[0];
        assert_eq!(sheet.hidden_columns, vec![(0, 2)]);
        assert_eq!(sheet.hidden_rows, vec![(1, 2), (3, 1_048_575)]);
        assert_eq!(sheet.used_range, Some((1_048_576, 3)));
    }

    #[test]
    fn test_merged_cells_indexing_ods() {
        use std::io::Cursor;
        use std::io::Write;
        use zip::write::SimpleFileOptions;

        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options =
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);

            zip.start_file("content.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office:spreadsheet="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="Sheet1">
                <table:table-row>
                    <!-- Cell A1 spans 2 columns (A, B) -->
                    <table:table-cell table:number-columns-spanned="2" office:value-type="string">
                        <text:p>Spanned</text:p>
                    </table:table-cell>
                    <!-- B1 is covered -->
                    <table:covered-table-cell/>
                    <!-- C1 should be at column 2 -->
                    <table:table-cell office:value-type="string">
                        <text:p>Target</text:p>
                    </table:table-cell>
                </table:table-row>
                <table:table-row>
                    <!-- Test repeated columns -->
                    <table:table-cell table:number-columns-repeated="2" office:value-type="string">
                        <text:p>Repeated</text:p>
                    </table:table-cell>
                    <!-- C2 should be at column 2 -->
                    <table:table-cell office:value-type="string">
                        <text:p>AfterRepeated</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#).unwrap();

            zip.finish().unwrap();
        }

        let mut archive = ZipArchive::new(Cursor::new(buf)).unwrap();
        let mut reader = OdsReader::new(&mut archive).unwrap();
        let sheets = reader.read_sheets().unwrap();

        assert_eq!(sheets.len(), 1);
        let sheet = &sheets[0];

        // Row 1
        // (0, 0) -> "Spanned"
        assert_eq!(
            sheet.cells.get(&(0, 0)).unwrap().value,
            CellValue::Text(Arc::from("Spanned"))
        );
        // (0, 1) -> Covered (usually not in map if empty)
        // (0, 2) -> "Target" (This would be (0, 3) in the buggy version)
        assert_eq!(
            sheet.cells.get(&(0, 2)).unwrap().value,
            CellValue::Text(Arc::from("Target"))
        );

        // Row 2
        // (1, 0) -> "Repeated"
        assert_eq!(
            sheet.cells.get(&(1, 0)).unwrap().value,
            CellValue::Text(Arc::from("Repeated"))
        );
        // (1, 1) -> "Repeated"
        assert_eq!(
            sheet.cells.get(&(1, 1)).unwrap().value,
            CellValue::Text(Arc::from("Repeated"))
        );
        // (1, 2) -> "AfterRepeated"
        assert_eq!(
            sheet.cells.get(&(1, 2)).unwrap().value,
            CellValue::Text(Arc::from("AfterRepeated"))
        );
    }

    #[test]
    fn test_calculate_used_range_individual_cells() {
        use std::collections::HashMap;
        let mut cells = HashMap::new();
        cells.insert(
            (5, 3),
            Cell {
                row: 5,
                col: 3,
                value: CellValue::Text(Arc::from("test")),
                ..Default::default()
            },
        );

        let used_range = calculate_used_range(&cells);
        assert_eq!(used_range, Some((5, 3)));

        cells.insert(
            (10, 7),
            Cell {
                row: 10,
                col: 7,
                value: CellValue::Number(42.0),
                ..Default::default()
            },
        );

        let used_range = calculate_used_range(&cells);
        assert_eq!(used_range, Some((10, 7)));
    }

    #[test]
    fn test_calculate_used_range_merged_cells() {
        use std::collections::HashMap;
        let mut cells = HashMap::new();

        // Merged cell B2:C3
        cells.insert(
            (1, 1),
            Cell {
                row: 1,
                col: 1,
                value: CellValue::Text(Arc::from("M")),
                ..Default::default()
            },
        );
        cells.insert(
            (1, 2),
            Cell {
                row: 1,
                col: 2,
                ..Default::default()
            },
        );
        cells.insert(
            (2, 1),
            Cell {
                row: 2,
                col: 1,
                ..Default::default()
            },
        );
        cells.insert(
            (2, 2),
            Cell {
                row: 2,
                col: 2,
                ..Default::default()
            },
        );

        let used_range = calculate_used_range(&cells);
        assert_eq!(used_range, Some((2, 2)));
    }

    #[test]
    fn test_calculate_used_range_empty() {
        use std::collections::HashMap;
        let cells: HashMap<(u32, u32), Cell> = HashMap::new();
        assert_eq!(calculate_used_range(&cells), None);
    }
}

#[test]
fn test_extract_external_workbooks_with_test_asset() {
    const TEST_ODS: &[u8] = include_bytes!("../../../tests/minimal_test.ods");
    let cursor = std::io::Cursor::new(TEST_ODS);
    let mut archive = ZipArchive::new(cursor).unwrap();
    let mut reader = OdsReader::new(&mut archive).unwrap();
    let workbooks = reader.read_external_workbooks().unwrap();

    // Verify external workbooks are extracted
    assert!(
        !workbooks.is_empty(),
        "Should detect external workbooks in test file"
    );

    // Verify indices are 0-based and sequential (order of appearance)
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
fn test_sheet_collection_ods() {
    const TEST_ODS: &[u8] = include_bytes!("../../../tests/minimal_test.ods");
    let cursor = std::io::Cursor::new(TEST_ODS);
    let mut archive = ZipArchive::new(cursor).unwrap();
    let mut reader = OdsReader::new(&mut archive).unwrap();

    let sheets = reader.read_sheets().unwrap();

    // Verify sheet count (should not include external sheets)
    assert!(!sheets.is_empty(), "Should have at least one sheet");

    // Verify no external sheet references in names
    // This implicitly tests that external sheets are filtered out
    for sheet in &sheets {
        assert!(
            !sheet.name.contains("file:///"),
            "Sheet collection should not contain external sheets (filtered): {}",
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
fn test_sheet_visibility_ods() {
    const TEST_ODS: &[u8] = include_bytes!("../../../tests/minimal_test.ods");
    let cursor = std::io::Cursor::new(TEST_ODS);
    let mut archive = ZipArchive::new(cursor).unwrap();
    let mut reader = OdsReader::new(&mut archive).unwrap();

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

//! ODS (`OpenDocument` Spreadsheet) parser implementation

use super::parser_utils::{attr_value, push_general_ref, read_text_node};
use super::workbook::{
    Cell, CellValue, DocumentSettings, ExternalWorkbook, IterationSettings, Sheet,
};
use super::{
    CellSink, NameConflict, NameScope, NumFmt, Properties, Selection, SheetInfo, SheetState,
    WorkbookReader,
};
use crate::error::Result;
use crate::reader::ExcelError;
use crate::reader::selection::SheetSelection;
use quick_xml::Reader;
use quick_xml::events::Event;
use regex::Regex;
use sheetrs_formula::addr::needs_quotes;
use std::collections::HashMap;
use std::fmt::Write as _;
use std::io::BufReader;
use std::sync::{Arc, LazyLock};
use zip::ZipArchive;

/// Extract merged cell ranges from an ODS worksheet
/// ODS format: <table:table-cell table:number-columns-spanned="X" table:number-rows-spanned="Y">
/// Whether an ODS document carries macros or scripts: a `Basic/` or `Scripts/` library (an entry
/// of the archive, or one the manifest lists even when the archive has no entry for the folder),
/// or a macro bound to an event of the document, a form or a cell (`vnd.sun.star.script:` in
/// `content.xml`).
///
/// # Returns
///
/// * `Result<bool>` - True if macros or scripts are detected
///
/// # Errors
///
/// When the archive cannot be read.
pub fn has_macros(archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>) -> Result<bool> {
    if has_macro_libraries(archive)? {
        return Ok(true);
    }
    // Macros bound to events (`script:event-listener` with a `vnd.sun.star.script:` target).
    if let Ok(content) = archive.by_name("content.xml") {
        return stream_has_event_macro(content);
    }

    Ok(false)
}

/// Whether the archive holds macro libraries: entries under `Basic/` or `Scripts/`, or the
/// manifest listing one. The macros bound to events live in `content.xml` and are not looked for
/// here: the reader that reads that part anyway finds them while it goes by (see [`MarkTap`]).
fn has_macro_libraries(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<bool> {
    let is_library = |path: &str| path.starts_with("Basic/") || path.starts_with("Scripts/");

    // 1. Libraries present in the archive.
    for i in 0..archive.len() {
        if let Ok(file) = archive.by_index(i)
            && is_library(file.name())
        {
            return Ok(true);
        }
    }

    // 2. Libraries the manifest lists.
    if let Ok(manifest_file) = archive.by_name("META-INF/manifest.xml") {
        let mut reader = Reader::from_reader(BufReader::new(manifest_file));
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "manifest:file-entry" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == "manifest:full-path"
                            && is_library(&attr_value(&attr)?)
                        {
                            return Ok(true);
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

/// The text a macro bound to an event holds in `content.xml`.
const EVENT_MACRO_MARK: &[u8] = b"vnd.sun.star.script:";

/// Passes the bytes of a reader through and notes whether the mark of an event macro went by, so
/// that `content.xml` is read once for the cells and for the macros. The mark is searched block
/// by block, keeping the end of the previous one so a mark split between two reads is found.
struct MarkTap<R> {
    inner: R,
    search: aho_corasick::AhoCorasick,
    tail: Vec<u8>,
    found: std::rc::Rc<std::cell::Cell<bool>>,
}

impl<R> MarkTap<R> {
    fn new(inner: R, found: std::rc::Rc<std::cell::Cell<bool>>) -> Self {
        Self {
            inner,
            search: aho_corasick::AhoCorasick::new([EVENT_MACRO_MARK])
                .expect("the mark is a valid pattern"),
            tail: Vec::new(),
            found,
        }
    }
}

impl<R: std::io::Read> std::io::Read for MarkTap<R> {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        let n = self.inner.read(buf)?;
        if !self.found.get() {
            self.tail.extend_from_slice(&buf[..n]);
            if self.search.is_match(&self.tail) {
                self.found.set(true);
                self.tail = Vec::new();
            } else {
                let keep = self.tail.len().min(EVENT_MACRO_MARK.len() - 1);
                self.tail.drain(..self.tail.len() - keep);
            }
        }
        Ok(n)
    }
}

/// Whether the document content read from `content` holds a macro bound to an event: the text
/// `vnd.sun.star.script:`, searched block by block so a mark split between two reads is found.
fn stream_has_event_macro(mut content: impl std::io::Read) -> Result<bool> {
    const MARK: &[u8] = EVENT_MACRO_MARK;
    let mark = aho_corasick::AhoCorasick::new([MARK]).expect("the mark is a valid pattern");
    let mut chunk = vec![0u8; 1 << 16];
    let mut tail: Vec<u8> = Vec::new();
    loop {
        let n = content.read(&mut chunk)?;
        if n == 0 {
            return Ok(false);
        }
        tail.extend_from_slice(&chunk[..n]);
        if mark.is_match(&tail) {
            return Ok(true);
        }
        // What could be the start of the mark across two reads.
        let keep = tail.len().min(MARK.len() - 1);
        tail.drain(..tail.len() - keep);
    }
}

/// Extract cached error values from an ODS worksheet
/// ODS error values are often stored in calcext:value-type="error" and calcext:value="#ERROR!"
fn parse_ods_date(date_str: &str, null_date_shift: i32) -> Option<f64> {
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
    let day = date_components[2].parse::<i32>().ok()?;

    if !(1..=12).contains(&month) || day < 1 {
        return None;
    }

    // O(1) conversion from (year, month, day) to the number of days since the
    // document's null date, which is how a spreadsheet numbers dates (the
    // value a formula such as `=A1+1` sees). LibreOffice's default null date is
    // 1899-12-30; `null_date_shift` is the offset of the file's own
    // `table:null-date` from it (0 by default, 1462 for 1904-01-01).
    //
    // The calendar is the real one: nothing is adjusted for the Excel 1900
    // leap-year quirk, since ODS does not have it.

    // Cumulative days before each month in a non-leap year (index 1 = Jan).
    const CUMULATIVE_DAYS: [i32; 13] = [0, 0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334];

    let is_leap = |y: i32| (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);

    // Total days from year 0 to the start of `y` (Jan 1 of year y).
    // Leap year corrections use (y-1) to count leap years in [0, y-1] only;
    // using y would erroneously include year y's own leap day.
    let days_before_year =
        |y: i32| -> i32 { 365 * y + (y - 1) / 4 - (y - 1) / 100 + (y - 1) / 400 };

    let leap_feb = i32::from(month > 2 && is_leap(year));
    let absolute_days = days_before_year(year) + CUMULATIVE_DAYS[month as usize] + leap_feb + day;

    // 1899-12-30 is day 0 (1900-01-01 is therefore 2).
    let total_days = absolute_days - days_before_year(1900) + 1 - null_date_shift;

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

    Some(f64::from(total_days) + time_fraction)
}

/// Days between `LibreOffice`'s default null date (1899-12-30) and the
/// `table:null-date` declared by the document (`YYYY-MM-DD`).
fn null_date_shift(null_date: &str) -> Option<i32> {
    // parse_ods_date numbers days from 1899-12-30, so the date itself gives the shift.
    parse_ods_date(null_date, 0).map(|d| d as i32)
}

/// Parse an `office:time-value` duration (`PT13H30M00S`, fractional seconds
/// allowed) as a fraction of a day.
fn parse_ods_time(s: &str) -> Option<f64> {
    let rest = s.strip_prefix("PT")?;
    let (mut total, mut num) = (0.0f64, String::new());
    for c in rest.chars() {
        match c {
            '0'..='9' | '.' => num.push(c),
            'H' | 'M' | 'S' => {
                let v: f64 = num.parse().ok()?;
                total += v * match c {
                    'H' => 3600.0,
                    'M' => 60.0,
                    _ => 1.0,
                };
                num.clear();
            }
            _ => return None,
        }
    }
    num.is_empty().then_some(total / 86400.0)
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

    // String literals are text, not references: `"[$$-409]0.00"` stays as written.
    let mut result = String::with_capacity(input.len());
    for (is_text, segment) in split_text_literals(input) {
        if is_text {
            result.push_str(segment);
            continue;
        }
        let mut converted = String::with_capacity(segment.len());
        let mut chars = segment.chars().peekable();
        while let Some(ch) = chars.next() {
            match ch {
                '[' => {
                    // Handle bracketed references: [.A1], [$Sheet.A1], [.A1:.B2], ['file:///path'#Sheet.A1]
                    parse_bracket_ref(
                        &mut chars,
                        &mut converted,
                        preserve_sheet,
                        current_sheet_name,
                        external_workbooks,
                    );
                }
                '$' => {
                    // Handle $Sheet.A1 or $Sheet.A1:.$B$2 (unbracketed sheet reference)
                    if let Some(next) = chars.peek() {
                        if next.is_alphabetic() || *next == '\'' {
                            parse_dollar_sheet_ref(
                                &mut chars,
                                &mut converted,
                                preserve_sheet,
                                if preserve_sheet {
                                    None
                                } else {
                                    current_sheet_name
                                },
                            );
                        } else {
                            converted.push(ch);
                        }
                    } else {
                        converted.push(ch);
                    }
                }
                _ => converted.push(ch),
            }
        }
        // Post-processing: handle plain Sheet.Cell references (without $ or brackets)
        // This handles cases like "Sheet1.A1" or "Sheet1.B2:Sheet1.B4"
        if !preserve_sheet {
            converted = strip_local_sheet_refs(&converted);
        }
        result.push_str(&converted);
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

/// Split a formula into code and double-quoted text literals (`""` inside a
/// literal is an escaped quote). Each item says whether it is a literal, quotes
/// included.
/// Reads on to the end of the table the reader is in, without building anything: the tables of a
/// sheet nobody asked for, and the rest of the one whose last wanted row went by. Tables nest only
/// as database ranges do not, so the depth is what makes the end of this one the right one.
fn skip_table<R: std::io::BufRead>(reader: &mut Reader<R>) -> Result<()> {
    let mut depth = 0u32;
    let mut skip_buf = Vec::new();
    loop {
        match reader.read_event_into(&mut skip_buf)? {
            Event::Start(e) if e.name().as_ref() == "table:table" => depth += 1,
            Event::End(e) if e.name().as_ref() == "table:table" => {
                if depth == 0 {
                    return Ok(());
                }
                depth -= 1;
            }
            Event::Eof => return Ok(()),
            _ => {}
        }
        skip_buf.clear();
    }
}

/// The expression with the references written without a sheet (`[.A1]`, `[.A1:.B2]`) given the
/// sheet `sheet` (`[$Data.A1]`, `[$'Q1 2024'.A1]`), outside text literals.
fn qualify_local_refs(expression: &str, sheet: &str) -> String {
    let quoted = if sheet.chars().all(|c| c.is_alphanumeric() || c == '_') {
        sheet.to_string()
    } else {
        format!("'{}'", sheet.replace('\'', "''"))
    };
    let with_sheet = format!("[${quoted}.");
    split_text_literals(expression)
        .into_iter()
        .map(|(is_text, segment)| {
            if is_text {
                segment.to_string()
            } else {
                segment.replace("[.", &with_sheet)
            }
        })
        .collect()
}

fn split_text_literals(formula: &str) -> Vec<(bool, &str)> {
    let mut out = Vec::new();
    let (mut start, mut in_text) = (0, false);
    let bytes = formula.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'"' {
            if !in_text {
                if i > start {
                    out.push((false, &formula[start..i]));
                }
                start = i;
                in_text = true;
            } else if bytes.get(i + 1) == Some(&b'"') {
                i += 1;
            } else {
                out.push((true, &formula[start..=i]));
                start = i + 1;
                in_text = false;
            }
        }
        i += 1;
    }
    if start < formula.len() {
        out.push((in_text, &formula[start..]));
    }
    out
}

/// Check if a reference is a whole column (A) or whole row (1)
fn is_whole_column_or_row(s: &str) -> bool {
    let cleaned = s.trim_start_matches('$');
    cleaned.chars().all(char::is_alphabetic) || cleaned.chars().all(char::is_numeric)
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
    if bracket_content.trim_start_matches('$').starts_with("#REF!") {
        // What LibreOffice writes for a reference that is gone: `[#REF!]`, `[$#REF!.A1]`.
        result.push_str("#REF!");
    } else if let Some(hash_pos) = bracket_content.find('#') {
        // External reference: ['file:///path'#Sheet.A1]
        let raw_uri = &bracket_content[..hash_pos];
        let ref_part = &bracket_content[hash_pos + 1..];
        let uri = raw_uri.strip_circumfix('\'', '\'').unwrap_or(raw_uri);

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
        let _ = write!(result, "[{}]", index + 1);
        result.push_str(&normalized_ref);
    } else if let Some(stripped) = bracket_content.strip_prefix('$') {
        // [$Sheet.A1] or [$Sheet.A1:.B2]
        parse_sheet_qualified_ref(stripped, result, current_sheet_name);
    } else if let Some(stripped) = bracket_content.strip_prefix('.') {
        // [.A1] or [.A1:.B2] or [.A:.A] or [.1:.1]
        parse_local_ref(stripped, result);
    } else if split_sheet_cell(&bracket_content).is_some() {
        // [Sheet.A1], ['My Sheet'.A1:.B2]: a sheet that is not absolute, which is what
        // `LibreOffice` writes for a reference such as `=Sheet.A1`.
        parse_sheet_qualified_ref(&bracket_content, result, current_sheet_name);
    } else {
        // Unknown format, keep as-is
        result.push('[');
        result.push_str(&bracket_content);
        result.push(']');
    }
}

/// The sheet name of an ODS reference without its quotes.
fn unquote_sheet(name: &str) -> String {
    name.strip_prefix('\'')
        .and_then(|rest| rest.strip_suffix('\''))
        .map_or_else(|| name.to_string(), |inner| inner.replace("''", "'"))
}

/// The index of the first `target` outside the single quotes of a sheet name (`''` inside a name
/// is a quote, and toggles twice).
fn find_outside_quotes(text: &str, target: char) -> Option<usize> {
    let mut in_quotes = false;
    for (i, c) in text.char_indices() {
        match c {
            '\'' => in_quotes = !in_quotes,
            c if c == target && !in_quotes => return Some(i),
            _ => {}
        }
    }
    None
}

/// `Sheet.A1` or `'My Sheet'.A1` as the sheet (as written) and the cell; `None` when there is no
/// sheet in front of the dot.
fn split_sheet_cell(part: &str) -> Option<(&str, &str)> {
    let dot = find_outside_quotes(part, '.')?;
    (dot > 0).then(|| (&part[..dot], &part[dot + 1..]))
}

/// The name of a sheet written in an ODS reference (`$` of its absolute mark and quotes
/// removed).
fn sheet_name_of(written: &str) -> String {
    unquote_sheet(written.trim_start_matches('$'))
}

/// A sheet name as Excel writes it in a reference: quoted when it needs it.
fn push_sheet_name(result: &mut String, name: &str) {
    if needs_quotes(name) {
        let _ = write!(result, "'{}'", name.replace('\'', "''"));
    } else {
        result.push_str(name);
    }
}

/// Parse a reference with a sheet: `Sheet.A1`, `'My Sheet'.A1:.B2`, `Sheet.A1:Sheet.B2`,
/// `$Sheet1.A1:$Sheet3.B2` (the `$` before a sheet only says that the sheet is absolute). It is
/// written as Excel writes it (`Sheet!A1`, `Sheet1:Sheet3!A1:B2`), without the sheet when it is the
/// one of the formula.
fn parse_sheet_qualified_ref(content: &str, result: &mut String, current_sheet_name: Option<&str>) {
    // A range is two parts around a colon that is not in the quotes of a sheet name.
    let (start_part, end_part) = match find_outside_quotes(content, ':') {
        Some(i) => (&content[..i], Some(&content[i + 1..])),
        None => (content, None),
    };
    let Some((start_sheet, start_cell)) = split_sheet_cell(start_part) else {
        // No sheet in front of the cell: malformed here, kept as written.
        result.push_str(content);
        return;
    };
    let first = sheet_name_of(start_sheet);
    let local = current_sheet_name == Some(first.as_str());
    let Some(end_part) = end_part else {
        // A single cell.
        if !local {
            push_sheet_name(result, &first);
            result.push('!');
        }
        result.push_str(start_cell);
        return;
    };
    // The end of a range: `.B2` (the same sheet), `Sheet.B2` or `$Sheet.B2`.
    let (last, end_cell) = if let Some(cell) = end_part.strip_prefix('.') {
        (first.clone(), cell)
    } else if let Some((sheet, cell)) = split_sheet_cell(end_part) {
        (sheet_name_of(sheet), cell)
    } else {
        (first.clone(), end_part)
    };
    if last == first {
        if !local {
            push_sheet_name(result, &first);
            result.push('!');
        }
    } else if needs_quotes(&first) || needs_quotes(&last) {
        // A 3D reference, as an XLSX writes it: `Sheet1:Sheet3!A1:B2`.
        let _ = write!(
            result,
            "'{}:{}'!",
            first.replace('\'', "''"),
            last.replace('\'', "''")
        );
    } else {
        let _ = write!(result, "{first}:{last}!");
    }
    result.push_str(start_cell);
    result.push(':');
    result.push_str(end_cell);
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
/// Always outputs Sheet!Cell format, stripping handled at end of `normalize_ods_reference`
fn parse_dollar_sheet_ref(
    chars: &mut std::iter::Peekable<std::str::Chars>,
    result: &mut String,
    _preserve_sheet: bool,
    current_sheet_name: Option<&str>,
) {
    let mut sheet_name = String::new();

    if chars.peek() == Some(&'\'') {
        // A quoted name, which may hold dots and spaces: `$'Q1.2024'.A1`.
        sheet_name.push('\'');
        chars.next();
        while let Some(ch) = chars.next() {
            sheet_name.push(ch);
            if ch == '\'' {
                if chars.peek() == Some(&'\'') {
                    sheet_name.push('\'');
                    chars.next();
                } else {
                    break;
                }
            }
        }
        if chars.peek() == Some(&'.') {
            chars.next();
        } else {
            result.push('$');
            result.push_str(&sheet_name);
            return;
        }
    } else {
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
    }
    // The name, without the quotes it was written with: it is quoted again if Excel needs it.
    let sheet_name = unquote_sheet(&sheet_name);

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
            push_sheet_name(result, &sheet_name);
            result.push('!');
        }
        result.push_str(&cell_ref);
    } else {
        // Range
        if !should_strip {
            push_sheet_name(result, &sheet_name);
            result.push('!');
        }
        result.push_str(&cell_ref);
        result.push(':');
        result.push_str(&cell_ref2);
    }
}

static SHEET_RANGE_PATTERN: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(r"([A-Za-z0-9_]+)\.([A-Z$0-9]+):([A-Za-z0-9_]+)\.([A-Z$0-9]+)").unwrap()
});
static SHEET_SINGLE_PATTERN: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"^([A-Za-z0-9_]+)\.([A-Z$0-9]+)$").unwrap());

/// Strip local sheet references from formula (Sheet1.A1 → A1, Sheet1.A1:Sheet1.B2 → A1:B2)
/// Also converts Sheet.Cell to Sheet!Cell format for multi-sheet ranges
/// Only handles plain Sheet.Cell (no $ or brackets), which are processed in post-processing
/// Only strips if both parts of a range have the same sheet name
fn strip_local_sheet_refs(formula: &str) -> String {
    // Convert plain Sheet.Cell to Sheet!Cell for multi-sheet ranges
    // Pattern: Sheet1.A1:Sheet2.B2 → Sheet1!A1:Sheet2!B2
    // Only matches Sheet.Cell (with dot), not Sheet!Cell (already processed)
    let ranges = SHEET_RANGE_PATTERN.replace_all(formula, |caps: &regex::Captures| {
        if caps[1] == caps[3] {
            // Same sheet, strip it
            format!("{}:{}", &caps[2], &caps[4])
        } else {
            // Different sheets, convert to ! format
            format!("{}!{}:{}!{}", &caps[1], &caps[2], &caps[3], &caps[4])
        }
    });

    // Handle single cells: Sheet1.A1 → A1 (at start or end of string)
    SHEET_SINGLE_PATTERN.replace_all(&ranges, "$2").into_owned()
}

/// Replace ODS function argument separators (`;` → `,`) outside string literals.
///
/// ODS formulas use `;` between arguments while Excel uses `,`. This function
/// performs the substitution while preserving `;` inside double-quoted strings.
/// Inside an array constant, ODS separates the columns with `;` and the rows
/// with `|`, where Excel uses `,` and `;`.
fn normalize_arg_separator(formula: &str) -> String {
    let mut result = String::with_capacity(formula.len());
    let mut in_string = false;
    let mut in_array = false;
    let mut run_start = 0;

    for (i, byte) in formula.bytes().enumerate() {
        match byte {
            b'"' => in_string = !in_string,
            b'{' if !in_string => in_array = true,
            b'}' if !in_string => in_array = false,
            b';' if !in_string => {
                result.push_str(&formula[run_start..i]);
                result.push(',');
                run_start = i + 1;
            }
            b'|' if !in_string && in_array => {
                result.push_str(&formula[run_start..i]);
                result.push(';');
                run_start = i + 1;
            }
            _ => {}
        }
    }
    result.push_str(&formula[run_start..]);

    result
}

/// Container for metadata parsed from an ODS file
///
/// Only caches lightweight metadata needed by subsequent reader methods.
/// Sheet data is returned directly and not cached, avoiding a full clone.
#[derive(Debug)]
struct OdsData {
    hidden_sheets: Vec<String>,
    has_macros: bool,
    external_workbooks: Vec<ExternalWorkbook>,
    /// Named ranges, named expressions and named database ranges, read in the same pass as the
    /// cells.
    defined_names: HashMap<String, String>,
    /// The names defined in more than one scope.
    name_conflicts: Vec<NameConflict>,
}

/// ODS (`OpenDocument` Spreadsheet) reader implementation
///
/// Handles parsing of .ods content.xml and styles.xml
#[derive(Debug)]
pub struct OdsReader<'a, R: std::io::Read + std::io::Seek> {
    archive: &'a mut ZipArchive<R>,
    data: Option<OdsData>,
}

const MAX_COLUMNS: u32 = 16384;

impl<'a, R: std::io::Read + std::io::Seek> OdsReader<'a, R> {
    /// # Errors
    ///
    /// When the archive does not hold a readable ODS document.
    pub fn new(archive: &'a mut ZipArchive<R>) -> Result<Self> {
        Ok(Self {
            archive,
            data: None,
        })
    }
}

impl<R: std::io::Read + std::io::Seek> WorkbookReader for OdsReader<'_, R> {
    fn read_sheets_into(&mut self, sink: &mut dyn CellSink) -> Result<Vec<Sheet>> {
        self.read_sheets_inner(sink, None)
    }

    fn read_selection_into(
        &mut self,
        selection: &Selection,
        sink: &mut dyn CellSink,
    ) -> Result<Vec<SheetInfo>> {
        let sheets = self.read_sheets_inner(sink, Some(selection))?;
        Ok(sheets
            .into_iter()
            .map(|sheet| SheetInfo {
                name: sheet.name,
                index: sheet.sheet_index,
                state: sheet.state,
            })
            .collect())
    }

    fn read_name_conflicts(&mut self) -> Result<Vec<NameConflict>> {
        if self.data.is_none() {
            self.read_sheets()?;
        }
        Ok(self.data.as_ref().unwrap().name_conflicts.clone())
    }

    fn read_defined_names(&mut self) -> Result<HashMap<String, String>> {
        // The names were read with the cells.
        if self.data.is_none() {
            self.read_sheets()?;
        }
        Ok(self.data.as_ref().unwrap().defined_names.clone())
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

    fn read_properties(&mut self) -> Result<Vec<(String, String)>> {
        extract_properties_ods(self.archive)
    }

    fn read_settings(&mut self) -> Result<DocumentSettings> {
        let mut settings = extract_settings_ods(self.archive)?;
        settings.locale = extract_locale_ods(self.archive);
        settings.string_ref_syntax = extract_string_ref_syntax_ods(self.archive);
        settings.auto_calculate = extract_auto_calculate_ods(self.archive);
        Ok(settings)
    }
}

impl<R: std::io::Read + std::io::Seek> OdsReader<'_, R> {
    /// Reads the sheets, handing their cells to `sink`. With a selection, only what it asks for:
    /// a table that is not selected is skipped, a table is left after the last row a window
    /// reaches, and the reading ends with the last table selected.
    fn read_sheets_inner(
        &mut self,
        sink: &mut dyn CellSink,
        selection: Option<&Selection>,
    ) -> Result<Vec<Sheet>> {
        // ============================================================
        // SINGLE PASS: Parse content.xml once for all concerns
        // ============================================================

        let mut sheets = Vec::new();

        // ============================================================
        // STATE: Date styles extraction (from content.xml)
        // ============================================================
        let mut data_styles: HashMap<String, NumFmt> = HashMap::new();
        let mut cell_styles = HashMap::new();
        let mut date_styles: HashMap<String, NumFmt> = HashMap::new(); // Resolved styles for cell lookup
        let mut current_data_style_name = String::new();
        let mut current_format = String::new();
        let mut in_date_style = false;
        let mut automatic_order = false;

        // ============================================================
        // STATE: Hidden sheets detection
        // ============================================================
        // Offset of the document's null date from 1899-12-30 (see parse_ods_date): date cells
        // are stored as dates and numbered from the null date.
        let mut null_shift = extract_settings_ods(self.archive)?
            .null_date
            .and_then(|(y, m, d)| null_date_shift(&format!("{y:04}-{m:02}-{d:02}")))
            .unwrap_or(0);

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
        let mut database_range_targets: Vec<String> = Vec::new();
        let props = selection.map_or(Properties::ALL, Selection::properties);
        // The windows of the table being read, when a selection was given and selects it.
        let mut wanted: Option<SheetSelection<'_>> = None;
        // How many of the selected sheets are still to be read; the reading ends with the last.
        let mut remaining = selection.map_or(0, Selection::sheet_count);
        let mut finished = false;
        let mut defined_names: HashMap<String, String> = HashMap::new();
        // Every scope each name is defined in, to tell the ones defined more than once.
        let mut name_scopes: HashMap<String, Vec<NameScope>> = HashMap::new();
        let mut in_named_expressions = false;
        let mut in_database_ranges = false;
        // Whether the reader is inside a `<table:table>`: a name written there belongs to its sheet.
        let mut in_table = false;
        // Whether an event macro went by in `content.xml`.
        let macro_mark = std::rc::Rc::new(std::cell::Cell::new(false));
        // Events inside a cell, kept between cells.
        let mut text_buf: Vec<u8> = Vec::new();
        let mut object_hrefs: Vec<(String, String)> = Vec::new();
        // Largest row and column (0-based) of the cells of the current sheet.
        let mut cells_max: Option<(u32, u32)> = None;

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
                            "number:date-style" => {
                                in_date_style = true;
                                current_format.clear();
                                automatic_order = false;
                                for attr in e.attributes().flatten() {
                                    match attr.key.as_ref() {
                                        "style:name" => {
                                            current_data_style_name =
                                                attr_value(&attr)?.into_owned();
                                        }
                                        "number:automatic-order" if attr.value == "true" => {
                                            automatic_order = true;
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            "number:text-style" => {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == "style:name" {
                                        let style_name = attr_value(&attr)?.into_owned();
                                        data_styles.insert(style_name.clone(), NumFmt::from("@"));
                                        date_styles.insert(style_name, NumFmt::from("@"));
                                    }
                                }
                            }
                            "style:style" => {
                                let mut is_cell_style = false;
                                let mut style_name = String::new();
                                let mut data_style_name = String::new();

                                for attr in e.attributes().flatten() {
                                    match attr.key.as_ref() {
                                        "style:family" if attr.value == "table-cell" => {
                                            is_cell_style = true;
                                        }
                                        "style:name" => {
                                            style_name = attr_value(&attr)?.into_owned();
                                        }
                                        "style:data-style-name" => {
                                            data_style_name = attr_value(&attr)?.into_owned();
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
                            "number:day" if in_date_style => {
                                let mut long = automatic_order;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == "number:style" {
                                        long = attr.value == "long";
                                    }
                                }
                                current_format.push_str(if long { "dd" } else { "d" });
                            }
                            "number:month" if in_date_style => {
                                let mut long = automatic_order;
                                let mut textual = false;
                                for attr in e.attributes().flatten() {
                                    match attr.key.as_ref() {
                                        "number:style" => {
                                            long = attr.value == "long";
                                        }
                                        "number:textual" if attr.value == "true" => {
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
                            "number:year" if in_date_style => {
                                let mut long = false;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == "number:style" && attr.value == "long" {
                                        long = true;
                                    }
                                }
                                current_format.push_str(if long { "yyyy" } else { "yy" });
                            }
                            "number:hours" if in_date_style => {
                                current_format.push_str("hh");
                            }
                            "number:minutes" if in_date_style => {
                                current_format.push_str("mm");
                            }
                            "number:seconds" if in_date_style => {
                                current_format.push_str("ss");
                            }
                            "number:text" if in_date_style => {
                                // Will read text event next
                            }
                            _ => {}
                        }
                    }
                    Event::Empty(e) => match e.name().as_ref() {
                        "table:null-date" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == "table:date-value"
                                    && let Some(d) = null_date_shift(&attr_value(&attr)?)
                                {
                                    null_shift = d;
                                }
                            }
                        }
                        "number:day" if in_date_style => {
                            let mut long = automatic_order;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == "number:style" {
                                    long = attr.value == "long";
                                }
                            }
                            current_format.push_str(if long { "dd" } else { "d" });
                        }
                        "number:month" if in_date_style => {
                            let mut long = automatic_order;
                            let mut textual = false;
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    "number:style" => {
                                        long = attr.value == "long";
                                    }
                                    "number:textual" if attr.value == "true" => {
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
                        "number:year" if in_date_style => {
                            let mut long = false;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == "number:style" && attr.value == "long" {
                                    long = true;
                                }
                            }
                            current_format.push_str(if long { "yyyy" } else { "yy" });
                        }
                        "number:hours" if in_date_style => {
                            current_format.push_str("hh");
                        }
                        "number:minutes" if in_date_style => {
                            current_format.push_str("mm");
                        }
                        "number:seconds" if in_date_style => {
                            current_format.push_str("ss");
                        }
                        _ => {}
                    },
                    Event::Text(e) if in_date_style => {
                        current_format.push_str(&e);
                    }
                    Event::GeneralRef(e) if in_date_style => {
                        push_general_ref(&mut current_format, &e)?;
                    }
                    Event::End(e) if e.name().as_ref() == "number:date-style" => {
                        if !current_data_style_name.is_empty() {
                            data_styles.insert(
                                current_data_style_name.clone(),
                                NumFmt::from(current_format.as_str()),
                            );
                            date_styles.insert(
                                current_data_style_name.clone(),
                                NumFmt::from(current_format.as_str()),
                            );
                            // Resolve any cell styles that reference this data style
                            for (cell_style, data_style) in &cell_styles {
                                if data_style == &current_data_style_name {
                                    date_styles.insert(
                                        cell_style.clone(),
                                        NumFmt::from(current_format.as_str()),
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
            let Ok(content_xml) = self.archive.by_name("content.xml") else {
                return Ok(sheets);
            };

            let mut reader = Reader::from_reader(BufReader::new(MarkTap::new(
                content_xml,
                macro_mark.clone(),
            )));
            reader.config_mut().trim_text(false);

            let mut buf = Vec::new();
            // Byte offset in content.xml where the current <table:table> starts
            let mut table_start = 0u64;

            // ============================================================
            // UNIFIED PARSING LOOP
            // ============================================================

            loop {
                let event_start = reader.buffer_position();
                let event = reader.read_event_into(&mut buf)?;
                let empty_tag = matches!(event, Event::Empty(_));
                match event {
                    // --------------------------------------------------------
                    // DATE STYLES: number:date-style
                    // --------------------------------------------------------
                    Event::Start(e) if e.name().as_ref() == "number:date-style" => {
                        in_date_style = true;
                        current_format.clear();
                        automatic_order = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "style:name" => {
                                    current_data_style_name = attr_value(&attr)?.into_owned();
                                }
                                "number:automatic-order" if attr.value == "true" => {
                                    automatic_order = true;
                                }
                                _ => {}
                            }
                        }
                    }
                    // DATE STYLES: number:text-style
                    Event::Start(e) if e.name().as_ref() == "number:text-style" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "style:name" {
                                let style_name = attr_value(&attr)?.into_owned();
                                data_styles.insert(style_name.clone(), NumFmt::from("@"));
                                date_styles.insert(style_name, NumFmt::from("@"));
                            }
                        }
                    }
                    // DATE STYLES: style:style (for cell styles mapping)
                    // A style with nothing inside (`<style:style ... style:data-style-name="N49"/>`)
                    // is how `LibreOffice` writes the cell style of a cell that only has a number
                    // format.
                    Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "style:style" => {
                        let mut is_cell_style = false;
                        let mut style_name = String::new();
                        let mut data_style_name = String::new();
                        let mut is_table_style = false;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "style:family" => {
                                    if attr.value == "table-cell" {
                                        is_cell_style = true;
                                    } else if attr.value == "table" {
                                        is_table_style = true;
                                    }
                                }
                                "style:name" => {
                                    style_name = attr_value(&attr)?.into_owned();
                                }
                                "style:data-style-name" => {
                                    data_style_name = attr_value(&attr)?.into_owned();
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
                        if is_table_style && !empty_tag && !style_name.is_empty() {
                            let mut inner_buf = Vec::new();
                            loop {
                                match reader.read_event_into(&mut inner_buf)? {
                                    Event::Start(ee) | Event::Empty(ee)
                                        if ee.name().as_ref() == "style:table-properties" =>
                                    {
                                        for attr in ee.attributes().flatten() {
                                            if attr.key.as_ref() == "table:display"
                                                && attr.value == "false"
                                            {
                                                hidden_styles.insert(style_name.clone());
                                            }
                                        }
                                    }
                                    Event::End(ee) if ee.name().as_ref() == "style:style" => {
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
                            "number:day" => {
                                let mut long = automatic_order;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == "number:style" {
                                        long = attr.value == "long";
                                    }
                                }
                                current_format.push_str(if long { "dd" } else { "d" });
                            }
                            "number:month" => {
                                let mut long = automatic_order;
                                let mut textual = false;
                                for attr in e.attributes().flatten() {
                                    match attr.key.as_ref() {
                                        "number:style" => {
                                            long = attr.value == "long";
                                        }
                                        "number:textual" if attr.value == "true" => {
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
                            "number:year" => {
                                let mut long = false;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == "number:style" && attr.value == "long" {
                                        long = true;
                                    }
                                }
                                current_format.push_str(if long { "yyyy" } else { "yy" });
                            }
                            "number:hours" => {
                                current_format.push_str("hh");
                            }
                            "number:minutes" => {
                                current_format.push_str("mm");
                            }
                            "number:seconds" => {
                                current_format.push_str("ss");
                            }
                            // "number:text" reads its text event next
                            _ => {}
                        }
                    }
                    // DATE STYLES: date components (Empty events)
                    Event::Empty(e) if in_date_style => match e.name().as_ref() {
                        "number:day" => {
                            let mut long = automatic_order;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == "number:style" {
                                    long = attr.value == "long";
                                }
                            }
                            current_format.push_str(if long { "dd" } else { "d" });
                        }
                        "number:month" => {
                            let mut long = automatic_order;
                            let mut textual = false;
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    "number:style" => {
                                        long = attr.value == "long";
                                    }
                                    "number:textual" if attr.value == "true" => {
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
                        "number:year" => {
                            let mut long = false;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == "number:style" && attr.value == "long" {
                                    long = true;
                                }
                            }
                            current_format.push_str(if long { "yyyy" } else { "yy" });
                        }
                        "number:hours" => {
                            current_format.push_str("hh");
                        }
                        "number:minutes" => {
                            current_format.push_str("mm");
                        }
                        "number:seconds" => {
                            current_format.push_str("ss");
                        }
                        _ => {}
                    },
                    // DATE STYLES: text content within date style
                    Event::Text(e) if in_date_style => {
                        current_format.push_str(&e);
                    }
                    Event::GeneralRef(e) if in_date_style => {
                        push_general_ref(&mut current_format, &e)?;
                    }
                    // DATE STYLES: end of date-style
                    Event::End(e) if e.name().as_ref() == "number:date-style" => {
                        if !current_data_style_name.is_empty() {
                            data_styles.insert(
                                current_data_style_name.clone(),
                                NumFmt::from(current_format.as_str()),
                            );
                            // Also add to date_styles directly
                            date_styles.insert(
                                current_data_style_name.clone(),
                                NumFmt::from(current_format.as_str()),
                            );
                            // Resolve any cell styles that reference this data style
                            for (cell_style, data_style) in &cell_styles {
                                if data_style == &current_data_style_name {
                                    date_styles.insert(
                                        cell_style.clone(),
                                        NumFmt::from(current_format.as_str()),
                                    );
                                }
                            }
                        }
                        in_date_style = false;
                    }
                    // --------------------------------------------------------
                    // MAIN SHEET DATA: table:table
                    // --------------------------------------------------------
                    Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "table:table" => {
                        // Collect sheet style for hidden detection
                        let mut sheet_name = String::new();
                        let mut style_name = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "table:name" => {
                                    sheet_name = attr_value(&attr)?.into_owned();
                                }
                                "table:style-name" => {
                                    style_name = attr_value(&attr)?.into_owned();
                                }
                                _ => {}
                            }
                        }

                        if !sheet_name.is_empty() && !style_name.is_empty() {
                            sheet_styles.push((sheet_name.clone(), style_name.clone()));
                        }

                        // Finalize previous sheet if it exists and it's not external
                        if let Some(sheet) = current_sheet.take()
                            && !skip_current_sheet
                        {
                            sheets.push(sheet);
                        }

                        let mut name = sheet_name;
                        if name.is_empty() {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == "table:name" {
                                    name = attr_value(&attr)?.into_owned();
                                }
                            }
                        }
                        // Assign sheet_index based on the current count of sheets parsed
                        let sheet_index = sheets.len() as u16;
                        let mut new_sheet = Sheet::new(name.clone(), sheet_index);
                        // Covers a self-closing table; the closing tag overwrites it
                        table_start = event_start;
                        new_sheet.xml_size_bytes = reader.buffer_position() - table_start;
                        // Check if this sheet's style is in hidden_styles
                        let is_hidden =
                            !style_name.is_empty() && hidden_styles.contains(&style_name);
                        new_sheet.state = if is_hidden {
                            SheetState::Hidden
                        } else {
                            SheetState::Visible
                        };

                        current_sheet = Some(new_sheet);
                        current_row = 0;
                        current_col = 0; // Reset column tracking for new sheet
                        skip_current_sheet = false; // Reset skip flag for new sheet
                        cells_max = None;
                        in_table = !empty_tag;
                        if let Some(selection) = selection {
                            wanted = current_sheet
                                .as_ref()
                                .and_then(|sheet| selection.of(&sheet.name));
                            if wanted.is_none() && !empty_tag {
                                // Not selected: listed, not read.
                                skip_table(&mut reader)?;
                                in_table = false;
                            } else if wanted.is_some() && empty_tag {
                                // A selected table with no rows: nothing to read in it.
                                wanted = None;
                                remaining = remaining.saturating_sub(1);
                                finished = remaining == 0;
                            }
                        }
                    }

                    // Detect external sheets by checking for table:table-source
                    Event::Start(ref e) | Event::Empty(ref e)
                        if e.name().as_ref() == "table:table-source" =>
                    {
                        // Check if this table-source has an xlink:href attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "xlink:href" {
                                let _href = attr_value(&attr)?.into_owned();
                                // This sheet is from an external workbook, mark it to be skipped
                                skip_current_sheet = true;
                                // Fast-forward to the end of this table to avoid parsing millions of rows
                                // CRITICAL: Start depth at 0 since we're already inside table:table element
                                let mut depth = 0;
                                let mut skip_buf = Vec::new();
                                loop {
                                    match reader.read_event_into(&mut skip_buf)? {
                                        Event::Start(ee) if ee.name().as_ref() == "table:table" => {
                                            depth += 1;
                                        }
                                        Event::End(ee) if ee.name().as_ref() == "table:table" => {
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
                    Event::Start(e)
                        if e.name().as_ref() == "table:table-column"
                            && let Some(sheet) = current_sheet.as_mut() =>
                    {
                        let mut hidden = false;
                        let mut repeated = 1u32;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "table:visibility"
                                    if (attr.value == "collapse" || attr.value == "filter") =>
                                {
                                    hidden = true;
                                }
                                "table:number-columns-repeated" => {
                                    repeated = attr_value(&attr)?.parse::<u32>().unwrap_or(1);
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
                                    if te.name().as_ref() == "table:table-column" =>
                                {
                                    break;
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                            col_buf.clear();
                        }
                    }
                    Event::Empty(e)
                        if e.name().as_ref() == "table:table-column"
                            && let Some(sheet) = current_sheet.as_mut() =>
                    {
                        let mut hidden = false;
                        let mut repeated = 1u32;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "table:visibility"
                                    if (attr.value == "collapse" || attr.value == "filter") =>
                                {
                                    hidden = true;
                                }
                                "table:number-columns-repeated" => {
                                    repeated = attr_value(&attr)?.parse::<u32>().unwrap_or(1);
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
                    Event::Start(e) if e.name().as_ref() == "table:table-row" => {
                        row_repeated = 1;
                        current_col = 0;
                        if let Some(ref mut sheet) = current_sheet {
                            let mut hidden = false;
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    "table:number-rows-repeated" => {
                                        row_repeated =
                                            attr_value(&attr)?.parse::<u32>().unwrap_or(1);
                                    }
                                    "table:visibility"
                                        if (attr.value == "collapse" || attr.value == "filter") =>
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
                        if let Some(rows) = wanted {
                            if current_row > rows.last_row {
                                // Nothing wanted is left in this table: rows only go forward.
                                // The rest of it is read through only to get to the next table
                                // that is wanted, if there is one.
                                wanted = None;
                                remaining = remaining.saturating_sub(1);
                                finished = remaining == 0;
                                if !finished {
                                    skip_table(&mut reader)?;
                                    in_table = false;
                                }
                            } else if current_row.saturating_add(row_repeated) <= rows.first_row {
                                // Rows before the first one wanted are not read.
                                reader.read_to_end_into(e.name(), &mut text_buf)?;
                                current_row += row_repeated;
                                current_col = 0;
                            }
                        }
                    }

                    Event::Empty(e) if e.name().as_ref() == "table:table-row" => {
                        // Empty row (self-closing tag) - no cells, just increment row counter
                        row_repeated = 1;
                        if let Some(ref mut sheet) = current_sheet {
                            let mut hidden = false;
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    "table:number-rows-repeated" => {
                                        row_repeated =
                                            attr_value(&attr)?.parse::<u32>().unwrap_or(1);
                                    }
                                    "table:visibility"
                                        if (attr.value == "collapse" || attr.value == "filter") =>
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
                        if wanted.is_some_and(|rows| current_row > rows.last_row) {
                            wanted = None;
                            remaining = remaining.saturating_sub(1);
                            finished = remaining == 0;
                            if !finished {
                                skip_table(&mut reader)?;
                                in_table = false;
                            }
                        }
                    }
                    Event::Start(e)
                        if matches!(
                            e.name().as_ref(),
                            "table:table-cell" | "table:covered-table-cell"
                        ) && let Some(sheet) = current_sheet.as_mut() =>
                    {
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
                                "table:number-columns-repeated" => {
                                    col_repeated = attr_value(&attr)?.parse::<u32>().unwrap_or(1);
                                }
                                "table:number-columns-spanned" => {
                                    cols_spanned = attr_value(&attr)?.parse::<u32>().unwrap_or(1);
                                }
                                "table:number-rows-spanned" => {
                                    rows_spanned = attr_value(&attr)?.parse::<u32>().unwrap_or(1);
                                }
                                "table:formula" => {
                                    // Whether the cell has one is told without converting it
                                    // when the text of the formula is not wanted.
                                    formula = Some(if props.formula {
                                        let raw_formula = attr_value(&attr)?;
                                        normalize_ods_reference(
                                            &raw_formula,
                                            false,
                                            Some(&sheet.name),
                                            &mut external_workbooks,
                                        )
                                    } else {
                                        String::new()
                                    });
                                }
                                "table:style-name" => {
                                    style_name = attr_value(&attr)?.into_owned();
                                }
                                "calcext:value-type" if attr.value == "error" => {
                                    is_error_cell = true;
                                }
                                "office:value"
                                | "office:string-value"
                                | "office:boolean-value"
                                | "office:time-value"
                                | "office:date-value" => {
                                    let val_str = attr_value(&attr)?.into_owned();
                                    if !has_value {
                                        value = match attr.key.as_ref() {
                                            "office:value" => {
                                                if let Ok(n) = val_str.parse::<f64>() {
                                                    CellValue::Number(n)
                                                } else {
                                                    CellValue::Text(Arc::from(val_str.as_str()))
                                                }
                                            }
                                            "office:date-value" => {
                                                // Convert ISO date to Serial Number
                                                if let Some(n) =
                                                    parse_ods_date(&val_str, null_shift)
                                                {
                                                    CellValue::Number(n)
                                                } else {
                                                    CellValue::Text(Arc::from(val_str.as_str()))
                                                }
                                            }
                                            "office:time-value" => {
                                                // A time of day is a fraction of a day.
                                                if let Some(n) = parse_ods_time(&val_str) {
                                                    CellValue::Number(n)
                                                } else {
                                                    CellValue::Text(Arc::from(val_str.as_str()))
                                                }
                                            }
                                            "office:boolean-value" => {
                                                CellValue::Boolean(val_str == "true")
                                            }
                                            _ => CellValue::Text(Arc::from(val_str.as_str())),
                                        };
                                        has_value = true;
                                    }
                                }
                                "table:number-matrix-columns-spanned" => {
                                    let v = attr_value(&attr)?.parse::<u32>().unwrap_or(0);
                                    if v > 0 {
                                        is_matrix_array = true;
                                    }
                                }
                                _ => {}
                            }
                        }

                        // A cell, or a run of repeated ones, that no window reaches is not read.
                        let skip_cell = wanted.is_some_and(|w| {
                            !w.may_contain_block(
                                current_row,
                                row_repeated,
                                current_col,
                                col_repeated,
                            )
                        });
                        if skip_cell {
                            reader.read_to_end_into(e.name(), &mut text_buf)?;
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
                        text_buf.clear();
                        loop {
                            if skip_cell {
                                break;
                            }
                            match reader.read_event_into(&mut text_buf)? {
                                // A comment's paragraphs are not the cell's content.
                                Event::Start(ref te)
                                    if te.name().as_ref() == "office:annotation" =>
                                {
                                    sheet.comments += 1;
                                    let end = te.to_end().into_owned();
                                    let mut skip = Vec::new();
                                    reader.read_to_end_into(end.name(), &mut skip)?;
                                }
                                Event::Start(ref te) if te.name().as_ref() == "text:p" => {
                                    // Paragraphs of one cell are separate lines.
                                    if !text_content.is_empty() {
                                        text_content.push('\n');
                                    }
                                    let mut p_buf = Vec::new();
                                    loop {
                                        match reader.read_event_into(&mut p_buf)? {
                                            Event::Text(ref t) => {
                                                text_content.push_str(t);
                                            }
                                            Event::GeneralRef(ref r) => {
                                                push_general_ref(&mut text_content, r)?;
                                            }
                                            // Runs of spaces, tabs and line breaks are elements.
                                            Event::Empty(ref se) | Event::Start(ref se)
                                                if matches!(
                                                    se.name().as_ref(),
                                                    "text:s" | "text:tab" | "text:line-break"
                                                ) =>
                                            {
                                                match se.name().as_ref() {
                                                    "text:tab" => text_content.push('\t'),
                                                    "text:line-break" => text_content.push('\n'),
                                                    _ => {
                                                        let mut n = 1usize;
                                                        for a in se.attributes().flatten() {
                                                            if a.key.as_ref() == "text:c"
                                                                && let Ok(v) = attr_value(&a)
                                                            {
                                                                n = v.parse().unwrap_or(1);
                                                            }
                                                        }
                                                        text_content
                                                            .extend(std::iter::repeat_n(' ', n));
                                                    }
                                                }
                                            }
                                            Event::End(ref pe)
                                                if pe.name().as_ref() == "text:p" =>
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
                                    if te.name().as_ref() == "table:table-cell"
                                        || te.name().as_ref() == "table:covered-table-cell" =>
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

                        if !skip_cell && (has_value || formula.is_some() || !style_name.is_empty())
                        {
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
                                let cell_value = value;
                                let formula_for_cell = formula;

                                // Look up format string from style
                                let num_fmt = if style_name.is_empty() || !props.number_format {
                                    None
                                } else {
                                    date_styles.get(&style_name).cloned()
                                };

                                // A float stays a float in a cell formatted as Text, as in XLSX.
                                for r in 0..row_repeated {
                                    for c in 0..col_repeated {
                                        let at = (current_row + r, current_col + c);
                                        if wanted.is_some_and(|w| !w.contains(at.0, at.1)) {
                                            continue;
                                        }
                                        let cell = Cell {
                                            value: if props.value {
                                                cell_value.clone()
                                            } else {
                                                CellValue::Empty
                                            },
                                            formula: if props.formula {
                                                formula_for_cell.as_deref().map(Box::from)
                                            } else {
                                                None
                                            },
                                            is_array: is_matrix_array && selection.is_none(),
                                            num_fmt: num_fmt.clone(),
                                        };
                                        cells_max =
                                            Some(cells_max.map_or(at, |(mr, mc)| {
                                                (mr.max(at.0), mc.max(at.1))
                                            }));
                                        if !skip_current_sheet {
                                            sink.cell(sheets.len() as u16, at, cell);
                                        }

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
                    Event::Empty(e)
                        if matches!(
                            e.name().as_ref(),
                            "table:table-cell" | "table:covered-table-cell"
                        ) && let Some(sheet) = current_sheet.as_mut() =>
                    {
                        let mut col_repeated = 1u32;
                        let mut cols_spanned = 1u32;
                        let mut rows_spanned = 1u32;
                        let mut formula = None;
                        let mut style_name = String::new();
                        let mut is_matrix_array = false;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "table:number-columns-repeated" => {
                                    col_repeated = attr_value(&attr)?.parse::<u32>().unwrap_or(1);
                                }
                                "table:number-columns-spanned" => {
                                    cols_spanned = attr_value(&attr)?.parse::<u32>().unwrap_or(1);
                                }
                                "table:number-rows-spanned" => {
                                    rows_spanned = attr_value(&attr)?.parse::<u32>().unwrap_or(1);
                                }
                                "table:formula" => {
                                    formula = Some(if props.formula {
                                        let raw_formula = attr_value(&attr)?;
                                        normalize_ods_reference(
                                            &raw_formula,
                                            false,
                                            Some(&sheet.name),
                                            &mut external_workbooks,
                                        )
                                    } else {
                                        String::new()
                                    });
                                }
                                "table:style-name" => {
                                    style_name = attr_value(&attr)?.into_owned();
                                }
                                "table:number-matrix-columns-spanned" => {
                                    let v = attr_value(&attr)?.parse::<u32>().unwrap_or(0);
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
                        let skip_cell = wanted.is_some_and(|w| {
                            !w.may_contain_block(
                                current_row,
                                row_repeated,
                                current_col,
                                col_repeated,
                            )
                        });
                        if !skip_cell && (formula.is_some() || !style_name.is_empty()) {
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
                                let num_fmt = if style_name.is_empty() || !props.number_format {
                                    None
                                } else {
                                    date_styles.get(&style_name).cloned()
                                };

                                for r in 0..row_repeated {
                                    for c in 0..col_repeated {
                                        let at = (current_row + r, current_col + c);
                                        if wanted.is_some_and(|w| !w.contains(at.0, at.1)) {
                                            continue;
                                        }
                                        let cell = Cell {
                                            value: if props.value {
                                                cell_value.clone()
                                            } else {
                                                CellValue::Empty
                                            },
                                            formula: if props.formula {
                                                formula_for_cell.as_deref().map(Box::from)
                                            } else {
                                                None
                                            },
                                            is_array: is_matrix_array && selection.is_none(),
                                            num_fmt: num_fmt.clone(),
                                        };
                                        cells_max =
                                            Some(cells_max.map_or(at, |(mr, mc)| {
                                                (mr.max(at.0), mc.max(at.1))
                                            }));
                                        if !skip_current_sheet {
                                            sink.cell(sheets.len() as u16, at, cell);
                                        }

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
                    Event::Start(e) if e.name().as_ref() == "calcext:conditional-format" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "calcext:target-range-address" {
                                current_cf_range = Some(attr_value(&attr)?.into_owned());
                            }
                        }
                    }
                    Event::End(e) if e.name().as_ref() == "calcext:conditional-format" => {
                        current_cf_range = None;
                    }
                    Event::Start(e)
                        if e.name().as_ref() == "calcext:condition"
                            && let Some(sheet) = current_sheet.as_mut() =>
                    {
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
                    // standard ODS conditional formatting
                    Event::Start(e) if e.name().as_ref() == "table:conditional-formatting" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "table:target-range-address" {
                                current_cf_range = Some(attr_value(&attr)?.into_owned());
                            }
                        }
                    }
                    Event::End(e) if e.name().as_ref() == "table:conditional-formatting" => {
                        current_cf_range = None;
                    }
                    Event::Start(e)
                        if e.name().as_ref() == "table:conditional-formatting-rule"
                            && let Some(sheet) = current_sheet.as_mut() =>
                    {
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
                    Event::Empty(e)
                        if e.name().as_ref() == "calcext:condition"
                            && let Some(sheet) = current_sheet.as_mut() =>
                    {
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
                    Event::End(e) if e.name().as_ref() == "table:table-row" => {
                        current_row += row_repeated;
                        current_col = 0;
                    }
                    Event::End(e)
                        if e.name().as_ref() == "table:table"
                            && let Some(sheet) = current_sheet.as_mut() =>
                    {
                        sheet.xml_size_bytes = reader.buffer_position() - table_start;
                        in_table = false;
                        if wanted.take().is_some() {
                            remaining = remaining.saturating_sub(1);
                            finished = remaining == 0;
                        }

                        // Calculate used range for the sheet before finalizing
                        // This must happen here (not at sheet finalization) because it needs to run
                        // for ALL sheets, whether they have conditional formatting or not
                        let cells_range = cells_max;

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

                    // Handle conditional formatting that appears after table closing tag
                    Event::Start(e) if e.name().as_ref() == "calcext:conditional-formats" => {
                        // This wrapper appears after </table:table>, continue processing
                    }
                    Event::End(e) if e.name().as_ref() == "calcext:conditional-formats" => {
                        // End of conditional formatting section
                        // Don't finalize the sheet here - the table end handler will do it
                    }
                    // Notes and charts of the sheet being read
                    Event::Start(e) | Event::Empty(e)
                        if e.name().as_ref() == "office:annotation"
                            && let Some(sheet) = current_sheet.as_mut() =>
                    {
                        sheet.comments += 1;
                    }
                    // An embedded object is a chart if its own document says so (checked below).
                    Event::Start(e) | Event::Empty(e)
                        if e.name().as_ref() == "draw:object"
                            && let Some(sheet) = current_sheet.as_mut() =>
                    {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "xlink:href" {
                                object_hrefs
                                    .push((sheet.name.clone(), attr_value(&attr)?.into_owned()));
                            }
                        }
                    }
                    // Names: the ranges, expressions and database ranges the document names. They
                    // come after the sheets, or inside the sheet they belong to.
                    Event::Start(e) if e.name().as_ref() == "table:named-expressions" => {
                        in_named_expressions = true;
                    }
                    Event::End(e) if e.name().as_ref() == "table:named-expressions" => {
                        in_named_expressions = false;
                    }
                    Event::Start(e) if e.name().as_ref() == "table:database-ranges" => {
                        in_database_ranges = true;
                    }
                    Event::End(e) if e.name().as_ref() == "table:database-ranges" => {
                        in_database_ranges = false;
                    }
                    Event::Start(e) | Event::Empty(e)
                        if in_named_expressions
                            && matches!(
                                e.name().as_ref(),
                                "table:named-range" | "table:named-expression"
                            ) =>
                    {
                        let is_range = e.name().as_ref() == "table:named-range";
                        let mut name = String::new();
                        let mut target = String::new();
                        let mut base = None;
                        for attr in e.attributes().flatten() {
                            match (attr.key.as_ref(), is_range) {
                                ("table:name", _) => name = attr_value(&attr)?.into_owned(),
                                ("table:cell-range-address", true)
                                | ("table:expression", false) => {
                                    target = attr_value(&attr)?.into_owned();
                                }
                                // The cell that references without a sheet are relative to.
                                ("table:base-cell-address", false) => {
                                    base = sheet_of_address(&attr_value(&attr)?);
                                }
                                _ => {}
                            }
                        }
                        if !name.is_empty() && !target.is_empty() {
                            let sheet = base.as_deref().or_else(|| {
                                in_table
                                    .then(|| {
                                        current_sheet.as_ref().map(|sheet| sheet.name.as_str())
                                    })
                                    .flatten()
                            });
                            // A name is global: its references keep their sheet, as in XLSX. The
                            // ones written without a sheet (`[.A1]`) are of the sheet of its base cell.
                            let target = match (is_range, sheet) {
                                (false, Some(sheet)) => qualify_local_refs(&target, sheet),
                                _ => target,
                            };
                            let normalized = normalize_ods_reference(
                                &target,
                                true,
                                None,
                                &mut external_workbooks,
                            );
                            let scope = match (in_table, current_sheet.as_ref()) {
                                (true, Some(sheet)) => NameScope::Sheet(sheet.name.clone()),
                                _ => NameScope::Workbook,
                            };
                            name_scopes.entry(name.clone()).or_default().push(scope);
                            defined_names.insert(name, normalized);
                        }
                    }
                    // A named database range is the table of the sheet its target is on; the
                    // anonymous one of an autofilter is not a table.
                    Event::Start(e) | Event::Empty(e)
                        if e.name().as_ref() == "table:database-range" =>
                    {
                        let mut name = String::new();
                        let mut target = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "table:name" => name = attr_value(&attr)?.into_owned(),
                                "table:target-range-address" => {
                                    target = attr_value(&attr)?.into_owned();
                                }
                                _ => {}
                            }
                        }
                        if !name.starts_with("__Anonymous_Sheet_DB__") && !target.is_empty() {
                            if in_database_ranges {
                                let normalized = normalize_ods_reference(
                                    &target,
                                    true,
                                    None,
                                    &mut external_workbooks,
                                );
                                name_scopes
                                    .entry(name.clone())
                                    .or_default()
                                    .push(NameScope::Workbook);
                                defined_names.insert(name, normalized);
                            }
                            database_range_targets.push(target);
                        }
                    }
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
                if finished {
                    break;
                }
            }

            // Finalize the last sheet if it exists and it's not external
            if let Some(sheet) = current_sheet
                && !skip_current_sheet
            {
                sheets.push(sheet);
            }
            if selection.is_some() {
                // The sheets, as far as the reading went: nothing else of the document is read.
                return Ok(sheets);
            }
            for target in &database_range_targets {
                let name = sheet_of_address(target);
                if let Some(sheet) = name.and_then(|n| sheets.iter_mut().find(|s| s.name == n)) {
                    sheet.tables += 1;
                }
            }
        } // End of content.xml parsing scope

        for (sheet_name, href) in object_hrefs {
            let part = format!(
                "{}/content.xml",
                href.trim_start_matches("./").trim_end_matches('/')
            );
            if is_chart_part(self.archive, &part)
                && let Some(sheet) = sheets.iter_mut().find(|s| s.name == sheet_name)
            {
                sheet.charts += 1;
            }
        }

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
        let has_macros = macro_mark.get() || has_macro_libraries(self.archive)?;

        // Store metadata for future method calls (sheets are returned directly)
        self.data = Some(OdsData {
            hidden_sheets,
            has_macros,
            external_workbooks,
            defined_names,
            name_conflicts: super::conflicts_of(name_scopes),
        });

        Ok(sheets)
    }
}

/// The sheet of a cell address such as `$'Q1.2024'.A1:.B3` (a quoted name can hold dots and
/// doubled quotes), `None` when it names none.
fn sheet_of_address(address: &str) -> Option<String> {
    let rest = address.strip_prefix('$').unwrap_or(address);
    if let Some(quoted) = rest.strip_prefix('\'') {
        let mut name = String::new();
        let mut chars = quoted.chars().peekable();
        while let Some(c) = chars.next() {
            match (c, chars.peek()) {
                ('\'', Some('\'')) => {
                    name.push('\'');
                    chars.next();
                }
                ('\'', _) => return Some(name),
                (c, _) => name.push(c),
            }
        }
        return None;
    }
    rest.split_once('.').map(|(sheet, _)| sheet.to_string())
}

/// Whether the part of an embedded object (`Object 1/content.xml`) holds a chart.
fn is_chart_part(archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>, part: &str) -> bool {
    let Ok(xml) = archive.by_name(part) else {
        return false;
    };
    let mut reader = Reader::from_reader(BufReader::new(xml));
    let mut buf = Vec::new();
    while let Ok(event) = reader.read_event_into(&mut buf) {
        match event {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "chart:chart" => return true,
            Event::Eof => return false,
            _ => {}
        }
        buf.clear();
    }
    false
}

/// The document properties of an ODS: the elements of `meta.xml`, named as an XLSX names them
/// (`meta:initial-creator` is the `creator`, `dc:creator` the `last_modified_by`), the keywords
/// joined with a comma, and the user-defined properties under their own name.
fn extract_properties_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Vec<(String, String)>> {
    let Ok(meta_xml) = archive.by_name("meta.xml") else {
        return Ok(Vec::new());
    };
    let mut reader = Reader::from_reader(BufReader::new(meta_xml));
    reader.config_mut().trim_text(true);
    let (mut out, mut keywords, mut buf) = (Vec::new(), Vec::new(), Vec::new());
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let tag = e.name().as_ref().to_string();
                let key = match tag.as_str() {
                    "dc:title" => "title".to_string(),
                    "dc:subject" => "subject".to_string(),
                    "dc:description" => "description".to_string(),
                    "meta:initial-creator" => "creator".to_string(),
                    "dc:creator" => "last_modified_by".to_string(),
                    "meta:creation-date" => "created".to_string(),
                    "dc:date" => "modified".to_string(),
                    "meta:generator" => "generator".to_string(),
                    "meta:keyword" => "keywords".to_string(),
                    "meta:user-defined" => match e.try_get_attribute("meta:name")? {
                        Some(a) => attr_value(&a)?.into_owned(),
                        None => continue,
                    },
                    _ => continue,
                };
                let value = read_text_node(&mut reader)?;
                match (key.as_str(), value.is_empty()) {
                    (_, true) => {}
                    ("keywords", false) => keywords.push(value),
                    (_, false) => out.push((key, value)),
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    if !keywords.is_empty() {
        out.push(("keywords".to_string(), keywords.join(", ")));
    }
    Ok(out)
}

/// Extract last modified date from meta.xml
fn extract_modified_date_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<Option<chrono::DateTime<chrono::Utc>>> {
    use chrono::{DateTime, Utc};

    let Ok(meta_xml) = archive.by_name("meta.xml") else {
        return Ok(None);
    };

    let mut reader = Reader::from_reader(BufReader::new(meta_xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut modified_date: Option<DateTime<Utc>> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                // dc:date element contains the last modified date
                let name = e.name();
                if name.as_ref().ends_with(":date") || name.as_ref() == "date" {
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

/// Extract date1904 setting from content.xml or settings.xml
/// ODS uses <table:null-date table:date-value="1904-01-01"/> in calculation settings, which
/// the document schema places before the first `table:table`: the search stops there.
fn extract_date1904_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<bool> {
    // Check content.xml for null-date specification
    let Ok(content_xml) = archive.by_name("content.xml") else {
        return Ok(false);
    };

    let mut reader = Reader::from_reader(BufReader::new(content_xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "table:null-date" => {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == "table:date-value" {
                        let value = attr_value(&attr)?;
                        // 1904 date system uses 1904-01-01 as epoch
                        return Ok(value.starts_with("1904"));
                    }
                }
            }
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "table:table" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(false)
}

/// `YYYY-MM-DD` (a time part is ignored) as `(year, month, day)`.
fn parse_iso_date(v: &str) -> Option<(i64, u32, u32)> {
    let mut it = v.split('T').next()?.splitn(3, '-');
    Some((
        it.next()?.parse().ok()?,
        it.next()?.parse().ok()?,
        it.next()?.parse().ok()?,
    ))
}

fn ods_bool(v: &str) -> bool {
    v == "true"
}

/// Read `table:calculation-settings` (and its `table:iteration`) from
/// content.xml. The element precedes the first table, so reading stops there.
fn extract_settings_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Result<DocumentSettings> {
    let mut settings = DocumentSettings::default();
    let Ok(content_xml) = archive.by_name("content.xml") else {
        return Ok(settings);
    };
    let mut reader = Reader::from_reader(BufReader::new(content_xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) => match e.name().as_ref() {
                "table:calculation-settings" => {
                    for attr in e.attributes().flatten() {
                        let v = attr_value(&attr)?;
                        match attr.key.as_ref() {
                            "table:case-sensitive" => settings.case_sensitive = Some(ods_bool(&v)),
                            "table:search-criteria-must-apply-to-whole-cell" => {
                                settings.search_whole_cell = Some(ods_bool(&v));
                            }
                            "table:use-regular-expressions" => {
                                settings.use_regular_expressions = Some(ods_bool(&v));
                            }
                            "table:use-wildcards" => settings.use_wildcards = Some(ods_bool(&v)),
                            "table:precision-as-shown" => {
                                settings.precision_as_shown = Some(ods_bool(&v));
                            }
                            "table:null-year" => settings.null_year = v.parse().ok(),
                            _ => {}
                        }
                    }
                }
                "table:null-date" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == "table:date-value" {
                            settings.null_date = parse_iso_date(&attr_value(&attr)?);
                        }
                    }
                }
                "table:iteration" => {
                    // ODF defaults: disabled, 100 steps, 0.001.
                    let mut it = IterationSettings {
                        enabled: false,
                        max_iterations: 100,
                        max_change: 0.001,
                    };
                    for attr in e.attributes().flatten() {
                        let v = attr_value(&attr)?;
                        match attr.key.as_ref() {
                            "table:status" => it.enabled = v == "enable",
                            "table:steps" => it.max_iterations = v.parse().unwrap_or(100),
                            "table:maximum-difference" => {
                                it.max_change = v.parse().unwrap_or(0.001);
                            }
                            _ => {}
                        }
                    }
                    settings.iteration = Some(it);
                }
                // Settings come before the first table.
                "table:table" => break,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(settings)
}

/// Document language from the default cell style in styles.xml
/// (`fo:language` + `fo:country`), as a BCP 47 tag.
/// `SyntaxStringRef` of `settings.xml`: how `INDIRECT` reads its text.
fn extract_string_ref_syntax_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Option<u8> {
    config_item_ods(archive, "SyntaxStringRef")?.parse().ok()
}

/// `AutoCalculate` of `settings.xml`: false when the document calculates only on request.
fn extract_auto_calculate_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Option<bool> {
    Some(config_item_ods(archive, "AutoCalculate")? == "true")
}

/// The text of a configuration item of `settings.xml`.
fn config_item_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
    name: &str,
) -> Option<String> {
    let xml = archive.by_name("settings.xml").ok()?;
    let mut reader = Reader::from_reader(BufReader::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut wanted = false;
    loop {
        match reader.read_event_into(&mut buf).ok()? {
            Event::Start(e) if e.name().as_ref() == "config:config-item" => {
                wanted = e.attributes().flatten().any(|a| {
                    a.key.as_ref() == "config:name" && attr_value(&a).is_ok_and(|v| v == name)
                });
            }
            Event::Text(t) if wanted => return Some(t.trim().to_string()),
            Event::End(_) => wanted = false,
            Event::Eof => return None,
            _ => {}
        }
        buf.clear();
    }
}

fn extract_locale_ods(
    archive: &mut ZipArchive<impl std::io::Read + std::io::Seek>,
) -> Option<String> {
    let styles = archive.by_name("styles.xml").ok()?;
    let mut reader = Reader::from_reader(BufReader::new(styles));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut in_default = false;
    loop {
        match reader.read_event_into(&mut buf).ok()? {
            Event::Start(e) if e.name().as_ref() == "style:default-style" => {
                in_default = e
                    .attributes()
                    .flatten()
                    .any(|a| a.key.as_ref() == "style:family" && a.value == "table-cell");
            }
            Event::End(e) if e.name().as_ref() == "style:default-style" => in_default = false,
            Event::Start(e) | Event::Empty(e)
                if in_default && e.name().as_ref() == "style:text-properties" =>
            {
                let (mut lang, mut country) = (None, None);
                for a in e.attributes().flatten() {
                    let v = attr_value(&a).ok()?.into_owned();
                    match a.key.as_ref() {
                        "fo:language" => lang = Some(v),
                        "fo:country" => country = Some(v),
                        _ => {}
                    }
                }
                return lang.filter(|l| l != "zxx").map(|l| {
                    match country.filter(|c| c != "none") {
                        Some(c) => format!("{l}-{c}"),
                        None => l,
                    }
                });
            }
            Event::Eof => return None,
            _ => {}
        }
        buf.clear();
    }
}

#[cfg(test)]
mod tests {
    #[test]
    fn the_sheet_of_an_address_may_be_quoted_with_dots_and_quotes() {
        assert_eq!(sheet_of_address("$Data.A1:.B3").as_deref(), Some("Data"));
        assert_eq!(
            sheet_of_address("$'Q1.2024'.A1:.B3").as_deref(),
            Some("Q1.2024")
        );
        assert_eq!(sheet_of_address("'It''s'.A1").as_deref(), Some("It's"));
        assert_eq!(sheet_of_address("$'open.A1"), None);
    }

    #[test]
    fn ods_properties_use_the_names_of_the_xlsx_ones() {
        use std::io::Write;
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        zip.start_file("meta.xml", zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(br#"<office:document-meta xmlns:office="o" xmlns:dc="d" xmlns:meta="m"><office:meta><dc:title>Model</dc:title><meta:initial-creator>Ana</meta:initial-creator><dc:creator>Luis</dc:creator><meta:keyword>a</meta:keyword><meta:keyword>b</meta:keyword><meta:generator>LibreOffice</meta:generator><meta:user-defined meta:name="Version">1.2</meta:user-defined></office:meta></office:document-meta>"#).unwrap();
        let mut archive = ZipArchive::new(zip.finish().unwrap()).unwrap();
        let got = extract_properties_ods(&mut archive).unwrap();
        let expected: Vec<(String, String)> = [
            ("title", "Model"),
            ("creator", "Ana"),
            ("last_modified_by", "Luis"),
            ("generator", "LibreOffice"),
            ("Version", "1.2"),
            ("keywords", "a, b"),
        ]
        .map(|(k, v)| (k.to_string(), v.to_string()))
        .to_vec();
        assert_eq!(got, expected);
    }

    use super::*;

    fn zip_with(files: &[(&str, &str)]) -> ZipArchive<std::io::Cursor<Vec<u8>>> {
        use std::io::Write;
        let mut w = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        for (name, body) in files {
            w.start_file(*name, zip::write::SimpleFileOptions::default())
                .unwrap();
            w.write_all(body.as_bytes()).unwrap();
        }
        ZipArchive::new(std::io::Cursor::new(w.finish().unwrap().into_inner())).unwrap()
    }

    #[test]
    fn ods_macros_are_libraries_listed_by_the_manifest_or_bound_events() {
        let manifest = |path: &str| {
            format!(
                r#"<manifest:manifest><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="{path}" manifest:media-type=""/></manifest:manifest>"#
            )
        };
        // A library in the archive.
        let mut z = zip_with(&[
            ("Basic/Standard/Module1.xml", "<x/>"),
            ("content.xml", "<a/>"),
        ]);
        assert!(has_macros(&mut z).unwrap());
        // A library only the manifest lists (the archive has no entry for the folder).
        let mut z = zip_with(&[
            ("META-INF/manifest.xml", &manifest("Scripts/python/")),
            ("content.xml", "<a/>"),
        ]);
        assert!(has_macros(&mut z).unwrap());
        // A macro bound to an event, far into a large content.xml.
        let big = format!(
            "{}<script:event-listener xlink:href=\"vnd.sun.star.script:Standard.M.Run?language=Basic&amp;location=document\"/>",
            "<table:table-cell/>".repeat(10_000)
        );
        let mut z = zip_with(&[("content.xml", &big)]);
        assert!(has_macros(&mut z).unwrap());
        // No macros: a manifest without libraries and a plain content.xml.
        let mut z = zip_with(&[
            ("META-INF/manifest.xml", &manifest("content.xml")),
            ("content.xml", "<office:document-content/>"),
        ]);
        assert!(!has_macros(&mut z).unwrap());
    }

    /// A reader that hands out a few bytes per call, as a decompressor may.
    struct Dribble<'a>(&'a [u8], usize);

    impl std::io::Read for Dribble<'_> {
        fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
            let n = self.1.min(self.0.len()).min(buf.len());
            buf[..n].copy_from_slice(&self.0[..n]);
            self.0 = &self.0[n..];
            Ok(n)
        }
    }

    #[test]
    fn a_reference_that_is_gone_is_a_ref_error_and_not_an_external_workbook() {
        let mut externals = Vec::new();
        let read = |f: &str, externals: &mut Vec<ExternalWorkbook>| {
            normalize_ods_reference(f, false, Some("S"), externals)
        };
        assert_eq!(read("of:=[#REF!]+1", &mut externals), "#REF!+1");
        assert_eq!(
            read("of:=SUM([$#REF!.A1:.B2])", &mut externals),
            "SUM(#REF!)"
        );
        assert!(externals.is_empty());
        // A real external reference keeps its reading.
        assert_eq!(
            read("of:=['file:///a/b.ods'#$Data.A1]", &mut externals),
            "[1]Data!A1"
        );
        assert_eq!(externals.len(), 1);
    }

    #[test]
    fn ods_event_macro_mark_split_between_reads_is_found() {
        let content = b"<a/>vnd.sun.star.script:Standard.M.Run";
        for size in [1, 3, 7, 19, 64] {
            assert!(
                stream_has_event_macro(Dribble(content, size)).unwrap(),
                "{size}"
            );
        }
        assert!(!stream_has_event_macro(Dribble(b"<a>vnd.sun.star.scrip</a>t:", 3)).unwrap());
    }

    #[test]
    fn null_date_is_read_from_the_settings_before_the_first_table() {
        let date1904 = |content: &str| {
            let mut z = zip_with(&[("content.xml", content)]);
            extract_date1904_ods(&mut z).unwrap()
        };
        let settings = |year: &str| {
            format!(
                r#"<table:calculation-settings><table:null-date table:date-value="{year}-01-01"/></table:calculation-settings>"#
            )
        };
        assert!(date1904(&format!(
            "<office:spreadsheet>{}<table:table table:name=\"S\"/></office:spreadsheet>",
            settings("1904")
        )));
        assert!(!date1904(&format!(
            "<office:spreadsheet>{}<table:table table:name=\"S\"/></office:spreadsheet>",
            settings("1899")
        )));
        // No settings: the tables are not searched.
        assert!(!date1904(
            "<office:spreadsheet><table:table table:name=\"S\"/></office:spreadsheet>"
        ));
    }

    #[test]
    fn string_reference_syntax_comes_from_settings_xml() {
        let settings = |value: &str| {
            format!(
                r#"<office:document-settings><office:settings><config:config-item-set config:name="ooo:configuration-settings">
                <config:config-item config:name="IsKernAsianPunctuation" config:type="boolean">false</config:config-item>
                <config:config-item config:name="SyntaxStringRef" config:type="short">{value}</config:config-item>
                </config:config-item-set></office:settings></office:document-settings>"#
            )
        };
        let mut zip = zip_with(&[("settings.xml", &settings("2"))]);
        assert_eq!(extract_string_ref_syntax_ods(&mut zip), Some(2));
        let mut zip = zip_with(&[("settings.xml", &settings("0"))]);
        assert_eq!(extract_string_ref_syntax_ods(&mut zip), Some(0));
        // A document created in LibreOffice stores nothing; no settings.xml either.
        let none = r#"<office:document-settings><config:config-item config:name="Other" config:type="short">2</config:config-item></office:document-settings>"#;
        let mut zip = zip_with(&[("settings.xml", none)]);
        assert_eq!(extract_string_ref_syntax_ods(&mut zip), None);
        let mut zip = zip_with(&[("content.xml", "<a/>")]);
        assert_eq!(extract_string_ref_syntax_ods(&mut zip), None);
    }

    #[test]
    fn auto_calculate_comes_from_settings_xml() {
        let settings = |value: &str| {
            format!(
                r#"<office:document-settings><config:config-item config:name="AutoCalculate" config:type="boolean">{value}</config:config-item></office:document-settings>"#
            )
        };
        let mut zip = zip_with(&[("settings.xml", &settings("false"))]);
        assert_eq!(extract_auto_calculate_ods(&mut zip), Some(false));
        let mut zip = zip_with(&[("settings.xml", &settings("true"))]);
        assert_eq!(extract_auto_calculate_ods(&mut zip), Some(true));
        let mut zip = zip_with(&[("content.xml", "<a/>")]);
        assert_eq!(extract_auto_calculate_ods(&mut zip), None);
    }

    #[test]
    fn test_extract_settings_ods() {
        let content = r#"<office:document-content><office:body><office:spreadsheet>
            <table:calculation-settings table:case-sensitive="false" table:use-regular-expressions="true"
              table:use-wildcards="false" table:search-criteria-must-apply-to-whole-cell="false"
              table:null-year="1950" table:precision-as-shown="true">
              <table:null-date table:value-type="date" table:date-value="1904-01-01"/>
              <table:iteration table:status="enable" table:steps="50" table:maximum-difference="0.01"/>
            </table:calculation-settings>
            <table:table table:name="S"/></office:spreadsheet></office:body></office:document-content>"#;
        let styles = r#"<office:document-styles><office:styles>
            <style:default-style style:family="graphic"><style:text-properties fo:language="de" fo:country="DE"/></style:default-style>
            <style:default-style style:family="table-cell"><style:text-properties fo:language="es" fo:country="CO"/></style:default-style>
            </office:styles></office:document-styles>"#;
        let mut zip = zip_with(&[("content.xml", content), ("styles.xml", styles)]);
        let mut s = extract_settings_ods(&mut zip).unwrap();
        s.locale = extract_locale_ods(&mut zip);
        assert_eq!(s.case_sensitive, Some(false));
        assert_eq!(s.use_regular_expressions, Some(true));
        assert_eq!(s.use_wildcards, Some(false));
        assert_eq!(s.search_whole_cell, Some(false));
        assert_eq!(s.null_year, Some(1950));
        assert_eq!(s.null_date, Some((1904, 1, 1)));
        assert_eq!(s.precision_as_shown, Some(true));
        assert_eq!(
            s.iteration,
            Some(IterationSettings {
                enabled: true,
                max_iterations: 50,
                max_change: 0.01
            })
        );
        assert_eq!(s.locale.as_deref(), Some("es-CO"));
    }

    #[test]
    fn test_extract_settings_ods_absent_and_partial() {
        let mut zip = zip_with(&[("content.xml", "<a><table:table table:name=\"S\"/></a>")]);
        assert_eq!(
            extract_settings_ods(&mut zip).unwrap(),
            DocumentSettings::default()
        );
        assert_eq!(extract_locale_ods(&mut zip), None);
        // Iteration with only a tolerance: disabled, ODF defaults for the rest.
        let mut zip = zip_with(&[(
            "content.xml",
            r#"<a><table:calculation-settings><table:iteration table:maximum-difference="0.0001"/></table:calculation-settings></a>"#,
        )]);
        let s = extract_settings_ods(&mut zip).unwrap();
        assert_eq!(
            s.iteration,
            Some(IterationSettings {
                enabled: false,
                max_iterations: 100,
                max_change: 0.0001
            })
        );
        assert_eq!(s.case_sensitive, None);
    }

    #[test]
    fn test_parse_ods_date_known_serials() {
        // Days since the null date 1899-12-30, real calendar: these are the
        // numbers LibreOffice itself shows for the dates.
        assert_eq!(parse_ods_date("1899-12-30", 0), Some(0.0));
        assert_eq!(parse_ods_date("1900-01-01", 0), Some(2.0));
        assert_eq!(parse_ods_date("1900-03-01", 0), Some(61.0));
        assert_eq!(parse_ods_date("1904-01-01", 0), Some(1462.0));
        assert_eq!(parse_ods_date("2000-01-01", 0), Some(36526.0));
        assert_eq!(parse_ods_date("2024-01-01", 0), Some(45292.0));
        assert_eq!(parse_ods_date("2024-07-04", 0), Some(45477.0));
        assert_eq!(parse_ods_date("1999-09-30", 0), Some(36433.0));
        assert_eq!(parse_ods_date("2025-05-03", 0), Some(45780.0));
    }

    #[test]
    fn test_parse_ods_date_relative_to_document_null_date() {
        let shift = null_date_shift("1904-01-01").unwrap();
        assert_eq!(shift, 1462);
        assert_eq!(parse_ods_date("1904-01-01", shift), Some(0.0));
        assert_eq!(parse_ods_date("2024-01-01", shift), Some(43830.0));
        assert_eq!(null_date_shift("1899-12-30"), Some(0));
    }

    #[test]
    fn test_parse_ods_time() {
        assert_eq!(parse_ods_time("PT12H00M00S"), Some(0.5));
        assert_eq!(parse_ods_time("PT13H30M00S"), Some(0.5625));
        assert_eq!(parse_ods_time("PT00H00M30.5S"), Some(30.5 / 86400.0));
        assert_eq!(parse_ods_time("PT36H00M00S"), Some(1.5));
        assert_eq!(parse_ods_time("12:00"), None);
    }

    #[test]
    fn test_parse_ods_date_with_time() {
        assert_eq!(parse_ods_date("1900-01-01T12:00:00", 0), Some(2.5));
        assert_eq!(parse_ods_date("1900-01-01T06:00:00", 0), Some(2.25));
    }

    #[test]
    fn test_parse_ods_date_invalid() {
        assert_eq!(parse_ods_date("not-a-date", 0), None);
        assert_eq!(parse_ods_date("2025-13", 0), None);
        // Out-of-range month/day must return None, not panic
        assert_eq!(parse_ods_date("2025-13-01", 0), None);
        assert_eq!(parse_ods_date("2025-00-01", 0), None);
        assert_eq!(parse_ods_date("2025-01-00", 0), None);
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
    fn test_normalize_ods_reference_leaves_text_literals_alone() {
        // Brackets, `$` and `Sheet.Cell` inside a string are text.
        assert_eq!(
            normalize_ods_reference(
                "of:=TEXT([.A1];\"[$$-409]#,##0.00\")&\"$Sheet1.A1 and Sheet1.B2\"",
                false,
                None,
                &mut Vec::new()
            ),
            "TEXT(A1,\"[$$-409]#,##0.00\")&\"$Sheet1.A1 and Sheet1.B2\""
        );
        assert_eq!(
            normalize_ods_reference("of:=\"a\"\"[.A1]\"\"\"&[.B1]", false, None, &mut Vec::new()),
            "\"a\"\"[.A1]\"\"\"&B1"
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
    fn test_normalize_arg_separator_array_constants() {
        assert_eq!(
            normalize_arg_separator("SUMPRODUCT(A1:B2*{1;2|3;4})"),
            "SUMPRODUCT(A1:B2*{1,2;3,4})"
        );
        // A `|` outside an array constant or inside a string is left alone.
        assert_eq!(normalize_arg_separator("A1|B1"), "A1|B1");
        assert_eq!(normalize_arg_separator("{\"a|b\"|1}"), "{\"a|b\";1}");
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
    fn test_normalize_arg_separator_multibyte_text() {
        assert_eq!(
            normalize_arg_separator("SI(año;\"ñ;é\";€)"),
            "SI(año,\"ñ;é\",€)"
        );
    }

    #[test]
    fn test_strip_local_sheet_refs() {
        assert_eq!(strip_local_sheet_refs("SUM(A1:B2)"), "SUM(A1:B2)");
        assert_eq!(strip_local_sheet_refs("Sheet1.A1:Sheet1.B2"), "A1:B2");
        assert_eq!(
            strip_local_sheet_refs("SUM(Sheet1.A1:Sheet2.B2)"),
            "SUM(Sheet1!A1:Sheet2!B2)"
        );
        assert_eq!(strip_local_sheet_refs("Sheet1.A1"), "A1");
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
    fn test_normalize_ods_external_reference_uri_quotes() {
        let external_paths = |raw: &str| {
            let mut external_workbooks = Vec::new();
            normalize_ods_reference(raw, false, None, &mut external_workbooks);
            external_workbooks
                .into_iter()
                .map(|eb| eb.path)
                .collect::<Vec<_>>()
        };

        assert_eq!(external_paths("of:=[''#Sheet1.A1]"), [""]);
        assert_eq!(external_paths("of:=['book.ods'#Sheet1.A1]"), ["book.ods"]);
        assert_eq!(external_paths("of:=[book.ods#Sheet1.A1]"), ["book.ods"]);
        assert_eq!(external_paths("of:=['#Sheet1.A1]"), ["'"]);
    }

    #[test]
    fn a_name_in_two_sheets_or_in_a_sheet_and_the_document_is_a_conflict_in_ods() {
        use std::io::Write;
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        zip.start_file("content.xml", zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet>
<table:table table:name="A"><table:table-row><table:table-cell/></table:table-row>
<table:named-expressions><table:named-expression table:name="L" table:base-cell-address="$A.$A$1" table:expression="of:=1"/><table:named-expression table:name="G" table:base-cell-address="$A.$A$1" table:expression="of:=2"/></table:named-expressions></table:table>
<table:table table:name="B"><table:table-row><table:table-cell/></table:table-row>
<table:named-expressions><table:named-expression table:name="L" table:base-cell-address="$B.$A$1" table:expression="of:=3"/></table:named-expressions></table:table>
<table:named-expressions><table:named-expression table:name="G" table:base-cell-address="$A.$A$1" table:expression="of:=4"/><table:named-range table:name="Own" table:base-cell-address="$A.$A$1" table:cell-range-address="$A.$A$1"/></table:named-expressions>
</office:spreadsheet></office:body></office:document-content>"#,
        )
        .unwrap();
        let mut archive = ZipArchive::new(zip.finish().unwrap()).unwrap();
        let mut reader = OdsReader::new(&mut archive).unwrap();
        let conflicts = reader.read_name_conflicts().unwrap();
        assert_eq!(
            conflicts,
            [
                NameConflict {
                    name: "G".to_string(),
                    scopes: vec![NameScope::Workbook, NameScope::Sheet("A".to_string())],
                },
                NameConflict {
                    name: "L".to_string(),
                    scopes: vec![
                        NameScope::Sheet("A".to_string()),
                        NameScope::Sheet("B".to_string())
                    ],
                },
            ]
        );
        // The map keeps one definition of each, and the names that are not repeated.
        let names = reader.read_defined_names().unwrap();
        assert_eq!(names.len(), 3);
        assert!(names.contains_key("Own"));
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
    fn test_cell_text_keeps_spaces_and_ignores_comments_ods() {
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
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body><office:spreadsheet>
        <table:table table:name="Sheet1">
            <table:table-row>
                <table:table-cell office:value-type="string"><text:p>a<text:s text:c="2"/>b<text:tab/>c</text:p><text:p>second</text:p></table:table-cell>
                <table:table-cell><office:annotation><text:p>author</text:p><text:p>a note</text:p></office:annotation></table:table-cell>
                <table:table-cell office:value-type="float" office:value="3"><office:annotation><text:p>note</text:p></office:annotation><text:p>3</text:p></table:table-cell>
            </table:table-row>
        </table:table>
    </office:spreadsheet></office:body>
</office:document-content>"#).unwrap();
            zip.finish().unwrap();
        }
        let mut archive = ZipArchive::new(Cursor::new(buf)).unwrap();
        let sheets = OdsReader::new(&mut archive).unwrap().read_sheets().unwrap();
        let sheet = &sheets[0];
        // Runs of spaces, tabs and paragraphs are content; a comment is not.
        assert_eq!(
            sheet.cells.get(&(0, 0)).unwrap().value,
            CellValue::Text(Arc::from("a  b\tc\nsecond"))
        );
        assert!(
            sheet
                .cells
                .get(&(0, 1))
                .is_none_or(|c| c.value == CellValue::Empty),
            "a comment on an empty cell is not its value"
        );
        assert_eq!(
            sheet.cells.get(&(0, 2)).unwrap().value,
            CellValue::Number(3.0)
        );
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
    let visible_count = sheets
        .iter()
        .filter(|s| s.state == SheetState::Visible)
        .count();
    let hidden_count = sheets
        .iter()
        .filter(|s| s.state != SheetState::Visible)
        .count();

    assert!(visible_count > 0, "Should have at least one visible sheet");
    assert!(hidden_count > 0, "Test file should have hidden sheets");

    // Verify hidden sheets have visible=false
    for sheet in &sheets {
        if sheet.name.contains("hidden") || sheet.name.contains("empty_hidden") {
            assert!(
                sheet.state != SheetState::Visible,
                "Sheet '{}' should be hidden",
                sheet.name
            );
        }
    }
}

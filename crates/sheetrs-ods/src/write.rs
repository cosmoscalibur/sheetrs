//! Writer of workbooks as ODS packages.
//!
//! A workbook is written as `LibreOffice` writes it: every table with its cells in row order,
//! repeated rows and cells for the gaps, and the names of the workbook after the tables.

use std::collections::HashMap;
use std::fmt::Write as _;
use std::io::{Cursor, Seek, Write};

use sheetrs_core::{
    CellPos, CellView, DateSystem, ErrorKind, FormulaGrammar, FormulaRef, NamedExpression,
    PackageFormat, SheetView, ValueRef, Visibility, Workbook,
};
use zip::CompressionMethod;
use zip::ZipArchive;
use zip::write::{SimpleFileOptions, ZipWriter};

use crate::Error;

const NAMESPACES: &str = r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0""#;
const MIME_TYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";
const FLUSH_BYTES: usize = 1 << 16;
/// The name of the style that hides a table.
const HIDDEN_STYLE: &str = "ta_hidden";

/// Writes `workbook` to `destination` as an ODS package.
///
/// A workbook read from an ODS package is written back into it: the entries other than
/// `content.xml` are copied as they were read, and `content.xml` is copied as it was read while
/// the workbook is as it was read (see [`sheetrs_core::Workbook::fingerprint`]), and is written
/// from the model otherwise.
///
/// # Errors
///
/// [`Error::Unsupported`] when the workbook has formulas that are not in the grammar of ODS,
/// and [`Error::Zip`] or [`Error::Io`] when the package cannot be written.
pub fn write<W: Write + Seek>(workbook: &Workbook, destination: W) -> Result<(), Error> {
    if workbook.formula_grammar() != FormulaGrammar::OpenFormula && !workbook.names().is_empty() {
        return Err(Error::Unsupported(
            "a name is not in the grammar of ODS".to_owned(),
        ));
    }
    let mut archive = ZipWriter::new(destination);
    let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    let package = workbook
        .opaque()
        .filter(|package| package.format == PackageFormat::Ods);
    // The media type goes first and is not compressed, which is how readers recognise the format.
    archive.start_file("mimetype", stored)?;
    archive.write_all(MIME_TYPE.as_bytes())?;
    if let Some(package) = package {
        for part in package.parts.iter().filter(|part| part.name != "mimetype") {
            archive.start_file(part.name.as_str(), deflated)?;
            archive.write_all(&part.bytes)?;
        }
        match &package.pristine {
            Some(pristine) if pristine.fingerprint == workbook.fingerprint() => {
                let mut source = ZipArchive::new(Cursor::new(pristine.entry.as_slice()))?;
                archive.raw_copy_file(source.by_index_raw(0)?)?;
            }
            _ => {
                archive.start_file("content.xml", deflated)?;
                write_content(&mut archive, workbook)?;
            }
        }
    } else {
        archive.start_file("META-INF/manifest.xml", deflated)?;
        archive.write_all(manifest().as_bytes())?;
        archive.start_file("content.xml", deflated)?;
        write_content(&mut archive, workbook)?;
    }
    archive.finish()?;
    Ok(())
}

fn manifest() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="{MIME_TYPE}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#
    )
}

fn write_content<W: Write>(destination: &mut W, workbook: &Workbook) -> Result<(), Error> {
    let mut out = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content {NAMESPACES} office:version="1.3">"#
    );
    out.push_str("<office:automatic-styles>");
    let _ = write!(
        out,
        r#"<style:style style:name="{HIDDEN_STYLE}" style:family="table"><style:table-properties table:display="false"/></style:style>"#
    );
    out.push_str("</office:automatic-styles><office:body><office:spreadsheet>");
    out.push_str(&calculation_settings(workbook.date_system()));
    for (index, sheet) in workbook.sheets().enumerate() {
        out.push_str(r#"<table:table table:name=""#);
        push_attribute(&mut out, sheet.name());
        out.push('"');
        if sheet.visibility() != Visibility::Visible {
            let _ = write!(out, r#" table:style-name="{HIDDEN_STYLE}""#);
        }
        out.push('>');
        write_rows(destination, &mut out, sheet, workbook.formula_grammar())?;
        out.push_str(&names(workbook.names(), Some(index), workbook));
        out.push_str("</table:table>");
    }
    out.push_str(&names(workbook.names(), None, workbook));
    out.push_str("</office:spreadsheet></office:body></office:document-content>");
    destination.write_all(out.as_bytes())?;
    Ok(())
}

/// The element that sets the date that serial number zero stands for, or nothing for the
/// default one.
fn calculation_settings(system: DateSystem) -> String {
    let zero = match system {
        DateSystem::Date1900 => return String::new(),
        DateSystem::Date1904 => -24_107,
        DateSystem::Custom { zero } => zero,
    };
    let (year, month, day) = civil_from_days(zero);
    format!(
        r#"<table:calculation-settings><table:null-date table:date-type="date" table:date-value="{year:04}-{month:02}-{day:02}"/></table:calculation-settings>"#
    )
}

/// The year, month and day of the day that is `days` after 1970-01-01.
fn civil_from_days(days: i64) -> (i64, i64, i64) {
    let shifted = days + 719_468;
    let era = shifted.div_euclid(146_097);
    let day_of_era = shifted.rem_euclid(146_097);
    let year_of_era =
        (day_of_era - day_of_era / 1_460 + day_of_era / 36_524 - day_of_era / 146_096) / 365;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let month_index = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * month_index + 2) / 5 + 1;
    let month = if month_index < 10 {
        month_index + 3
    } else {
        month_index - 9
    };
    let year = year_of_era + era * 400 + i64::from(month <= 2);
    (year, month, day)
}

/// Writes the rows of `sheet`, flushing `out` to `destination` as it grows.
fn write_rows<W: Write>(
    destination: &mut W,
    out: &mut String,
    sheet: SheetView<'_>,
    grammar: FormulaGrammar,
) -> Result<(), Error> {
    let members = sheet.array_members();
    let areas: HashMap<CellPos, (u32, u32)> = sheet
        .array_areas()
        .iter()
        .map(|(first, last)| (*first, (last.col - first.col + 1, last.row - first.row + 1)))
        .collect();
    let mut open_row: Option<u32> = None;
    let mut next_row = 0;
    let mut next_col = 0;
    for (pos, cell) in sheet.cells() {
        if open_row != Some(pos.row) {
            if open_row.is_some() {
                out.push_str("</table:table-row>");
            }
            if pos.row > next_row {
                let _ = write!(
                    out,
                    r#"<table:table-row table:number-rows-repeated="{}"/>"#,
                    pos.row - next_row
                );
            }
            out.push_str("<table:table-row>");
            open_row = Some(pos.row);
            next_row = pos.row + 1;
            next_col = 0;
        }
        if pos.col > next_col {
            let _ = write!(
                out,
                r#"<table:table-cell table:number-columns-repeated="{}"/>"#,
                pos.col - next_col
            );
        }
        next_col = pos.col + 1;
        let matrix = areas.get(&pos).copied();
        write_cell(out, sheet, &cell, grammar, matrix, members.contains(&pos))?;
        if out.len() >= FLUSH_BYTES {
            destination.write_all(out.as_bytes())?;
            out.clear();
        }
    }
    if open_row.is_some() {
        out.push_str("</table:table-row>");
    }
    Ok(())
}

fn write_cell(
    out: &mut String,
    sheet: SheetView<'_>,
    cell: &CellView<'_>,
    grammar: FormulaGrammar,
    matrix: Option<(u32, u32)>,
    covered: bool,
) -> Result<(), Error> {
    let text = match cell.formula {
        None => None,
        Some(FormulaRef::Text(text)) => Some(text),
        Some(FormulaRef::Shared { index }) => sheet
            .shared_anchor(index)
            .and_then(|anchor| sheet.get(anchor))
            .and_then(|anchor| match anchor.formula {
                Some(FormulaRef::Text(text)) => Some(text),
                _ => None,
            }),
    };
    if text.is_some() && grammar != FormulaGrammar::OpenFormula {
        return Err(Error::Unsupported(
            "a formula is not in the grammar of ODS".to_owned(),
        ));
    }
    let value = &cell.value;
    // The formula of a matrix is in its first cell, and the other cells it covers have none.
    let element = if covered {
        "table:covered-table-cell"
    } else {
        "table:table-cell"
    };
    let _ = write!(out, "<{element}");
    if covered {
        // Nothing but the value.
    } else if let Some(text) = text {
        out.push_str(r#" table:formula=""#);
        push_attribute(out, text);
        out.push('"');
        if let Some((columns, rows)) = matrix {
            let _ = write!(
                out,
                r#" table:number-matrix-columns-spanned="{columns}" table:number-matrix-rows-spanned="{rows}""#
            );
        }
    } else if let ValueRef::Error(kind) = value {
        // A constant error is not a value of a cell in ODS: it is the formula that is that error.
        let _ = write!(out, r#" table:formula="of:={}""#, error_text(kind));
    }
    match value {
        ValueRef::Empty => out.push_str("/>"),
        ValueRef::Number(number) => {
            let _ = write!(
                out,
                r#" office:value-type="float" office:value="{}"/>"#,
                number_text(*number)
            );
        }
        ValueRef::Bool(flag) => {
            let _ = write!(
                out,
                r#" office:value-type="boolean" office:boolean-value="{flag}"/>"#
            );
        }
        ValueRef::Error(kind) => {
            out.push_str(
                r#" office:value-type="string" office:string-value="" calcext:value-type="error"><text:p>"#,
            );
            push_inline(out, error_text(kind));
            let _ = write!(out, "</text:p></{element}>");
        }
        ValueRef::Text(text) => {
            out.push_str(r#" office:value-type="string" calcext:value-type="string">"#);
            push_paragraphs(out, text);
            let _ = write!(out, "</{element}>");
        }
    }
    Ok(())
}

/// The text of an error as `LibreOffice` shows it; the codes only it has are shown as `#VALUE!`.
fn error_text(kind: &ErrorKind) -> &str {
    match kind {
        ErrorKind::Other(_) => "#VALUE!",
        known => known.as_text(),
    }
}

/// Whether XML 1.0 allows `character` in a document, which `LibreOffice` does not write otherwise.
fn is_xml_character(character: char) -> bool {
    matches!(
        character,
        '\t' | '\n' | '\r' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
    )
}

/// Appends `text` as the content of an element, with the characters that XML does not allow
/// left out.
fn push_inline(out: &mut String, text: &str) {
    for character in text
        .chars()
        .filter(|character| is_xml_character(*character))
    {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            other => out.push(other),
        }
    }
}

/// The text of a number as `office:value` holds it: 15 significant digits, in full from `1E-14`
/// up to 2^53 and with an exponent of three digits outside, and `-0` for a negative zero.
fn number_text(number: f64) -> String {
    let Some(found) = sheetrs_io::fifteen_digits(number) else {
        return if number.is_sign_negative() { "-0" } else { "0" }.to_owned();
    };
    // Above 2^53 a whole number is no longer exact, and is written with an exponent.
    let exact = number.abs() < 9_007_199_254_740_992.0;
    if (-14..=15).contains(&found.exponent) && exact {
        found.plain(number)
    } else {
        found.scientific(3)
    }
}

/// Appends `text` as the paragraphs of a cell: a paragraph per line, runs of spaces as
/// `text:s` and tabs as `text:tab`.
fn push_paragraphs(out: &mut String, text: &str) {
    for line in text.split('\n') {
        out.push_str("<text:p>");
        let mut previous_space = true;
        let mut spaces = 0_usize;
        for character in line
            .chars()
            .filter(|character| is_xml_character(*character))
        {
            if character == ' ' && !previous_space {
                out.push(' ');
                previous_space = true;
                continue;
            }
            if character == ' ' {
                spaces += 1;
                continue;
            }
            flush_spaces(out, &mut spaces);
            previous_space = false;
            match character {
                '\t' => out.push_str("<text:tab/>"),
                '&' => out.push_str("&amp;"),
                '<' => out.push_str("&lt;"),
                '>' => out.push_str("&gt;"),
                '\r' => out.push_str("&#13;"),
                other => out.push(other),
            }
        }
        flush_spaces(out, &mut spaces);
        out.push_str("</text:p>");
    }
}

fn flush_spaces(out: &mut String, spaces: &mut usize) {
    match *spaces {
        0 => {}
        1 => out.push_str("<text:s/>"),
        count => {
            let _ = write!(out, r#"<text:s text:c="{count}"/>"#);
        }
    }
    *spaces = 0;
}

fn push_attribute(out: &mut String, text: &str) {
    for character in text
        .chars()
        .filter(|character| is_xml_character(*character))
    {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\n' => out.push_str("&#10;"),
            '\r' => out.push_str("&#13;"),
            '\t' => out.push_str("&#9;"),
            other => out.push(other),
        }
    }
}

/// The `table:named-expressions` element of the names in `scope`, or nothing when there are none.
/// A name whose expression is the address of a range is a named range; any other is a named
/// expression.
fn names(all: &[NamedExpression], scope: Option<usize>, workbook: &Workbook) -> String {
    let mut out = String::new();
    let base = workbook.sheets().next().map_or_else(
        || "$Sheet1.$A$1".to_owned(),
        |sheet| format!("${}.$A$1", sheet.name()),
    );
    for name in all.iter().filter(|name| name.scope == scope) {
        let mut element = String::new();
        let kind = if is_range_address(&name.expression) {
            "table:named-range"
        } else {
            "table:named-expression"
        };
        let attribute = if kind == "table:named-range" {
            "table:cell-range-address"
        } else {
            "table:expression"
        };
        let _ = write!(element, r#"<{kind} table:name=""#);
        push_attribute(&mut element, &name.name);
        element.push_str(r#"" table:base-cell-address=""#);
        push_attribute(&mut element, &base);
        let _ = write!(element, r#"" {attribute}=""#);
        push_attribute(&mut element, &name.expression);
        element.push_str(r#""/>"#);
        out.push_str(&element);
    }
    if out.is_empty() {
        out
    } else {
        format!("<table:named-expressions>{out}</table:named-expressions>")
    }
}

/// Whether `expression` is nothing but the address of a cell or a range, such as `$S1.$A$1:$B$2`.
fn is_range_address(expression: &str) -> bool {
    let cell = |text: &str| {
        let Some((_, address)) = text.rsplit_once('.') else {
            return false;
        };
        let address = address.replace('$', "");
        let letters = address.bytes().take_while(u8::is_ascii_alphabetic).count();
        letters > 0
            && address.len() > letters
            && address[letters..].bytes().all(|byte| byte.is_ascii_digit())
    };
    expression.starts_with('$')
        && expression.split(':').count() <= 2
        && expression.split(':').all(cell)
}

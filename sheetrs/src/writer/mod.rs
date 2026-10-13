//! Writer module for manipulating Excel/ODS files

mod ods_writer;
mod xlsx_writer;

pub use ods_writer::modify_workbook_ods;
pub use xlsx_writer::modify_workbook_xlsx;

use anyhow::{Result, bail};

use std::borrow::Cow;
use std::collections::HashSet;
use std::fmt;
use std::fs::File;
use std::io::{BufRead, Write};
use std::path::Path;
use std::str::FromStr;

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::reader::parser_utils::{attr_value, column_letters, column_letters_to_index};

/// Last 0-based column index of a worksheet (`XFD`).
const MAX_COL: u32 = 16_383;
/// Last 1-based row number of a worksheet.
const MAX_ROW: u32 = 1_048_576;

/// Struct used to define modifications to be applied to a workbook
#[derive(Debug, Default)]
pub struct WorkbookModifications {
    pub remove_sheets: Option<HashSet<String>>,
    pub remove_named_ranges: Option<HashSet<String>>,
    /// Replace formulas by their cached values in these cells.
    pub convert_to_values: Vec<ValueScope>,
}

/// Cells of one sheet whose formulas are replaced by their cached values.
///
/// Parsed from `Sheet`, `Sheet!B`, `Sheet!B:D`, `Sheet!3` or `Sheet!3:10`.
/// Quote a sheet name that contains `!` (`'R!D'!B:D`), doubling any `'`
/// inside it, as in a formula reference.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ValueScope {
    pub sheet: String,
    pub target: ValueTarget,
}

/// Part of a sheet covered by a [`ValueScope`]. Bounds are 0-based and
/// inclusive.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ValueTarget {
    Sheet,
    Columns(u32, u32),
    Rows(u32, u32),
}

impl ValueScope {
    /// Whether the 0-based cell `(row, col)` falls inside this scope.
    pub fn contains(&self, row: u32, col: u32) -> bool {
        match self.target {
            ValueTarget::Sheet => true,
            ValueTarget::Columns(first, last) => (first..=last).contains(&col),
            ValueTarget::Rows(first, last) => (first..=last).contains(&row),
        }
    }

    /// Whether the 0-based range from `first` to `last` (`(row, col)`, both
    /// inclusive) falls inside this scope.
    pub fn contains_range(&self, first: (u32, u32), last: (u32, u32)) -> bool {
        // A scope is a whole sheet, a column band or a row band, so holding
        // two opposite corners means holding the whole range.
        self.contains(first.0, first.1) && self.contains(last.0, last.1)
    }
}

impl FromStr for ValueScope {
    type Err = anyhow::Error;

    fn from_str(spec: &str) -> Result<Self> {
        let (sheet, target) = split_scope(spec)?;
        if sheet.is_empty() {
            bail!("missing sheet name in '{spec}'");
        }
        let target = match target {
            Some(target) => parse_target(target)
                .ok_or_else(|| anyhow::anyhow!("invalid columns or rows '{target}' in '{spec}'"))?,
            None => ValueTarget::Sheet,
        };
        Ok(Self { sheet, target })
    }
}

impl fmt::Display for ValueScope {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.sheet.contains(['!', '\'']) {
            write!(f, "'{}'", self.sheet.replace('\'', "''"))?;
        } else {
            f.write_str(&self.sheet)?;
        }
        match self.target {
            ValueTarget::Sheet => Ok(()),
            ValueTarget::Columns(first, last) => {
                write!(f, "!{}:{}", column_letters(first), column_letters(last))
            }
            ValueTarget::Rows(first, last) => write!(f, "!{}:{}", first + 1, last + 1),
        }
    }
}

/// Split a scope spec into its unquoted sheet name and optional target.
fn split_scope(spec: &str) -> Result<(String, Option<&str>)> {
    let Some(quoted) = spec.strip_prefix('\'') else {
        // An unquoted name keeps its `!` unless the text after the last `!`
        // is a valid target.
        return Ok(match spec.rsplit_once('!') {
            Some((sheet, target)) if parse_target(target).is_some() => {
                (sheet.to_string(), Some(target))
            }
            _ => (spec.to_string(), None),
        });
    };

    let mut sheet = String::new();
    let mut chars = quoted.char_indices();
    while let Some((i, ch)) = chars.next() {
        if ch != '\'' {
            sheet.push(ch);
            continue;
        }
        let rest = &quoted[i + 1..];
        if rest.starts_with('\'') {
            sheet.push('\'');
            chars.next();
            continue;
        }
        return match rest.strip_prefix('!') {
            Some(target) => Ok((sheet, Some(target))),
            None if rest.is_empty() => Ok((sheet, None)),
            None => bail!("unexpected '{rest}' after quoted sheet name in '{spec}'"),
        };
    }
    bail!("unterminated quoted sheet name in '{spec}'")
}

/// Parse `B`, `B:D`, `3` or `3:10` into 0-based inclusive bounds.
fn parse_target(target: &str) -> Option<ValueTarget> {
    let (first, last) = target.split_once(':').unwrap_or((target, target));
    if is_column(first) && is_column(last) {
        let first = column_letters_to_index(first).filter(|&c| c <= MAX_COL)?;
        let last = column_letters_to_index(last).filter(|&c| c <= MAX_COL)?;
        Some(ValueTarget::Columns(first.min(last), first.max(last)))
    } else {
        let first = parse_row(first)?;
        let last = parse_row(last)?;
        Some(ValueTarget::Rows(first.min(last), first.max(last)))
    }
}

fn is_column(part: &str) -> bool {
    // Longer than `XFD` could overflow the index computation.
    (1..=3).contains(&part.len()) && part.bytes().all(|b| b.is_ascii_alphabetic())
}

fn parse_row(part: &str) -> Option<u32> {
    if part.is_empty() || !part.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let row = part.parse::<u32>().ok()?;
    (1..=MAX_ROW).contains(&row).then(|| row - 1)
}

/// The unescaped value of the `key` attribute of an XML element.
fn attribute(element: &BytesStart, key: &str) -> Result<Option<String>> {
    Ok(element
        .try_get_attribute(key)?
        .map(|attr| attr_value(&attr).map(Cow::into_owned))
        .transpose()?)
}

/// Read the events of the element `name` up to its end tag, which is
/// consumed but not returned.
fn read_element_body<R: BufRead>(
    reader: &mut Reader<R>,
    name: &str,
) -> Result<Vec<Event<'static>>> {
    let mut body = Vec::new();
    let mut buf = Vec::new();
    let mut depth = 0u32;
    loop {
        let event = reader.read_event_into(&mut buf)?.into_owned();
        buf.clear();
        match &event {
            Event::Start(e) if e.name().as_ref() == name => depth += 1,
            Event::End(e) if e.name().as_ref() == name => {
                if depth == 0 {
                    return Ok(body);
                }
                depth -= 1;
            }
            Event::Eof => bail!("unexpected end of document inside <{name}>"),
            _ => {}
        }
        body.push(event);
    }
}

/// Write an element from its start tag and content, or as an empty element
/// when it has no content.
fn write_element<W: Write>(
    writer: &mut Writer<W>,
    start: BytesStart,
    content: Option<&[Event]>,
) -> Result<()> {
    match content {
        None => writer.write_event(Event::Empty(start))?,
        Some(content) => {
            let end = start.to_end().into_owned();
            writer.write_event(Event::Start(start))?;
            for event in content {
                writer.write_event(event.borrow())?;
            }
            writer.write_event(Event::End(end))?;
        }
    }
    Ok(())
}

/// Modify a workbook file (supports multiple operations)
pub fn modify_workbook<P: AsRef<Path>>(
    input_path: P,
    output_path: P,
    modifications: &WorkbookModifications,
) -> Result<()> {
    let input = input_path.as_ref();
    let output = output_path.as_ref();

    // Creating the output truncates it before the input is read.
    if output.exists() && output.canonicalize()? == input.canonicalize()? {
        bail!("Output file is the input file: {}", output.display());
    }
    let input_file = File::open(input)?;
    let output_file = File::create(output)?;

    // Determine file type by extension
    match input.extension().and_then(|s| s.to_str()) {
        Some("xlsx") => modify_workbook_xlsx(input_file, output_file, modifications),
        Some("ods") => modify_workbook_ods(input_file, output_file, modifications),
        _ => anyhow::bail!("Unsupported file format"),
    }
}

/// Legacy wrapper for removing sheets
pub fn remove_sheets<P: AsRef<Path>>(
    input_path: P,
    output_path: P,
    sheet_names: &[String],
) -> Result<()> {
    let modifications = WorkbookModifications {
        remove_sheets: Some(sheet_names.iter().cloned().collect()),
        ..Default::default()
    };
    modify_workbook(input_path, output_path, &modifications)
}

/// Legacy wrapper for removing named ranges
pub fn remove_named_ranges<P: AsRef<Path>>(
    input_path: P,
    output_path: P,
    range_names: &[String],
) -> Result<()> {
    let modifications = WorkbookModifications {
        remove_named_ranges: Some(range_names.iter().cloned().collect()),
        ..Default::default()
    };
    modify_workbook(input_path, output_path, &modifications)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn scope(spec: &str) -> ValueScope {
        spec.parse().unwrap()
    }

    #[test]
    fn test_parse_whole_sheet() {
        assert_eq!(
            scope("Data"),
            ValueScope {
                sheet: "Data".to_string(),
                target: ValueTarget::Sheet,
            }
        );
    }

    #[test]
    fn test_parse_columns_and_rows() {
        assert_eq!(scope("Data!B:D").target, ValueTarget::Columns(1, 3));
        assert_eq!(scope("Data!c").target, ValueTarget::Columns(2, 2));
        assert_eq!(scope("Data!D:B").target, ValueTarget::Columns(1, 3));
        assert_eq!(
            scope("Data!XFD").target,
            ValueTarget::Columns(MAX_COL, MAX_COL)
        );
        assert_eq!(scope("Data!3:10").target, ValueTarget::Rows(2, 9));
        assert_eq!(scope("Data!5").target, ValueTarget::Rows(4, 4));
    }

    #[test]
    fn test_parse_quoted_sheet() {
        let parsed = scope("'R!D ''x'''!A:A");
        assert_eq!(parsed.sheet, "R!D 'x'");
        assert_eq!(parsed.target, ValueTarget::Columns(0, 0));
        assert_eq!(scope("'My Sheet'").target, ValueTarget::Sheet);
    }

    #[test]
    fn test_unquoted_bang_without_target_is_sheet_name() {
        let parsed = scope("Q1!Totals");
        assert_eq!(parsed.sheet, "Q1!Totals");
        assert_eq!(parsed.target, ValueTarget::Sheet);
    }

    #[test]
    fn test_parse_rejects_invalid_specs() {
        for spec in [
            "",
            "!A:B",
            "'Data",
            "'Data'x",
            "'Data'!B3",
            "'Data'!0",
            "'Data'!XFE",
            "'Data'!AAAAAAAA",
            "'Data'!A:3",
            "'Data'!1048577",
        ] {
            assert!(spec.parse::<ValueScope>().is_err(), "{spec} should fail");
        }
    }

    #[test]
    fn test_display_round_trips() {
        for spec in ["Data", "Data!B:D", "Data!3:10", "'R!D ''x'''!A:A"] {
            assert_eq!(
                scope(spec).to_string().parse::<ValueScope>().unwrap(),
                scope(spec)
            );
        }
        assert_eq!(scope("Data!B").to_string(), "Data!B:B");
    }

    #[test]
    fn test_contains() {
        assert!(scope("Data").contains(7, 7));
        assert!(scope("Data!B:C").contains(100, 2));
        assert!(!scope("Data!B:C").contains(0, 3));
        assert!(scope("Data!2:3").contains(2, 50));
        assert!(!scope("Data!2:3").contains(0, 0));
    }
}

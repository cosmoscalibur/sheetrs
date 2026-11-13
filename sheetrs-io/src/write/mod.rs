//! The writers of XLSX and ODS: they rewrite a workbook as a stream, changing the parts the changes
//! touch and copying the rest as it is.
//!
//! This is the feature `write` of `sheetrs-io`. What decides the changes (the plan of an editor, a
//! search that replaces, a program that computed results) lives in other crates; what is here is
//! how a [`WorkbookModifications`] is applied to the archive of a workbook
//! ([`modify_workbook`], [`modify_workbook_bytes`], [`Book::modify`]), and the values and
//! constants a caller decided ([`CellWrites`]).

mod cells;
mod error;
pub(crate) mod formats;
mod ods_formats;
mod ods_structure;
mod ods_writer;
mod operation;
mod shapes;
pub(crate) mod structure;
#[cfg(test)]
mod workbooks;
mod xlsx_structure;
mod xlsx_writer;

pub use cells::{CellWrites, Constant};
pub use error::{Context, Error, Result, ScopeError};
pub use formats::{display_code, ods_value};
pub use ods_writer::modify_workbook_ods;
pub use operation::{Calculation, Operation, ReplaceSpec};
pub use structure::Bands;
pub use xlsx_writer::modify_workbook_xlsx;

use std::borrow::Cow;
use std::collections::HashSet;
use std::fmt;
use std::fs::File;
use std::io::{BufRead, Write};
use std::path::Path;
use std::str::FromStr;

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer};
use zip::write::SimpleFileOptions;

use crate::Book;
use crate::reader::Format;
use crate::reader::parser_utils::{attr_value, column_letters, column_letters_to_index};

/// Last 0-based column index of a worksheet (`XFD`).
const MAX_COL: u32 = 16_383;
/// Last 1-based row number of a worksheet.
const MAX_ROW: u32 = 1_048_576;

/// What a cell holds after a replacement, written in the notation of the file: a formula in Excel's
/// syntax for an XLSX and in `OpenFormula` (without the `of:=`) for an ODS.
#[derive(Debug, Clone, PartialEq)]
pub enum CellEdit {
    Clear,
    Text(String),
    Number(f64),
    Boolean(bool),
    Formula(String),
    /// A text `LibreOffice` could not make a formula of, or one that names what it does not have: the
    /// cell holds the text as a formula and the error `code` (`Err:511`; 525 is `#NAME?`, 524 `#REF!`).
    Invalid {
        formula: String,
        code: u16,
    },
    /// The cell keeps what it is (its formula, its array, its style) and the result it shows
    /// becomes this one.
    Cached(CachedResult),
    /// The note of the cell has this text; the cell itself is not touched.
    Note(String),
    /// A number or a boolean with the number format `LibreOffice` gives what was typed: the style of
    /// the cell stays, with this format.
    Formatted {
        edit: Box<CellEdit>,
        format: crate::EntryFormat,
        /// The text of `office:date-value` / `office:time-value`, in an ODS.
        ods_value: Option<String>,
        /// The text a number cell shows with the format, in the date system of the workbook, in an
        /// ODS.
        shown: Option<String>,
        /// The cell style of an ODS that has this format, once the writer made it.
        style: Option<String>,
    },
}

impl CellEdit {
    /// The edit under the format, if it has one.
    pub(crate) fn inner(&self) -> &Self {
        match self {
            Self::Formatted { edit, .. } => edit,
            other => other,
        }
    }
}

/// The result a cell shows, as a program computed it.
#[derive(Debug, Clone, PartialEq)]
pub enum CachedResult {
    Number(f64),
    Text(String),
    Boolean(bool),
    /// An error value written as the application shows it (`#DIV/0!`, `Err:522`).
    Error(String),
}

/// The cells of one sheet to rewrite, by `(row, column)`, 0-based.
pub type SheetEdits = std::collections::BTreeMap<(u32, u32), CellEdit>;

/// Struct used to define modifications to be applied to a workbook
#[derive(Debug, Default, Clone)]
pub struct WorkbookModifications {
    pub delete_sheets: Option<HashSet<String>>,
    pub delete_names: Option<HashSet<String>>,
    /// Replace formulas by their cached values in these cells.
    pub formula_to_value: Vec<ValueScope>,
    /// Delete these rows (`Sheet!3:10`), in the coordinates of the input.
    pub delete_rows: Vec<ValueScope>,
    /// Delete these columns (`Sheet!B:D`), in the coordinates of the input.
    pub delete_columns: Vec<ValueScope>,
    /// Cells whose content is replaced, by sheet name (the replacements of a pass, already
    /// decided against the input of the pass).
    pub replacements: std::collections::HashMap<String, SheetEdits>,
    /// Mark the workbook so the application that opens it first recalculates every formula.
    pub recalculate_on_open: bool,
    /// Take off the mark to recalculate when the workbook is opened, because the results the
    /// input holds were just calculated (an XLSX has the mark in `calcPr`; an ODS has none).
    pub clear_recalculation_mark: bool,
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
    #[must_use]
    pub fn contains(&self, row: u32, col: u32) -> bool {
        match self.target {
            ValueTarget::Sheet => true,
            ValueTarget::Columns(first, last) => (first..=last).contains(&col),
            ValueTarget::Rows(first, last) => (first..=last).contains(&row),
        }
    }

    /// Whether the 0-based range from `first` to `last` (`(row, col)`, both
    /// inclusive) falls inside this scope.
    #[must_use]
    pub fn contains_range(&self, first: (u32, u32), last: (u32, u32)) -> bool {
        // A scope is a whole sheet, a column band or a row band, so holding
        // two opposite corners means holding the whole range.
        self.contains(first.0, first.1) && self.contains(last.0, last.1)
    }
}

impl FromStr for ValueScope {
    type Err = ScopeError;

    fn from_str(spec: &str) -> Result<Self, ScopeError> {
        let (sheet, target) = split_scope(spec)?;
        if sheet.is_empty() {
            return Err(ScopeError::MissingSheet {
                spec: spec.to_string(),
            });
        }
        let target = match target {
            Some(target) => parse_target(target).ok_or_else(|| ScopeError::InvalidTarget {
                target: target.to_string(),
                spec: spec.to_string(),
            })?,
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

/// Rows of a sheet: `Sheet!3:10`, or `Sheet!7` for one. Bounds are 0-based and inclusive.
///
/// The type of the rows a deletion takes: one that is not a row range cannot be built, so no
/// operation has to check it later.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct RowRange {
    pub sheet: String,
    pub first: u32,
    pub last: u32,
}

/// Columns of a sheet: `Sheet!B:D`, or `Sheet!B` for one. Bounds are 0-based and inclusive.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ColumnRange {
    pub sheet: String,
    pub first: u32,
    pub last: u32,
}

impl FromStr for RowRange {
    type Err = ScopeError;

    fn from_str(spec: &str) -> Result<Self, ScopeError> {
        match spec.parse::<ValueScope>()? {
            ValueScope {
                sheet,
                target: ValueTarget::Rows(first, last),
            } => Ok(Self { sheet, first, last }),
            _ => Err(ScopeError::NotRows {
                spec: spec.to_string(),
            }),
        }
    }
}

impl FromStr for ColumnRange {
    type Err = ScopeError;

    fn from_str(spec: &str) -> Result<Self, ScopeError> {
        match spec.parse::<ValueScope>()? {
            ValueScope {
                sheet,
                target: ValueTarget::Columns(first, last),
            } => Ok(Self { sheet, first, last }),
            _ => Err(ScopeError::NotColumns {
                spec: spec.to_string(),
            }),
        }
    }
}

impl From<RowRange> for ValueScope {
    fn from(range: RowRange) -> Self {
        Self {
            sheet: range.sheet,
            target: ValueTarget::Rows(range.first, range.last),
        }
    }
}

impl From<ColumnRange> for ValueScope {
    fn from(range: ColumnRange) -> Self {
        Self {
            sheet: range.sheet,
            target: ValueTarget::Columns(range.first, range.last),
        }
    }
}

impl fmt::Display for RowRange {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        ValueScope::from(self.clone()).fmt(f)
    }
}

impl fmt::Display for ColumnRange {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        ValueScope::from(self.clone()).fmt(f)
    }
}

/// Split a scope spec into its unquoted sheet name and optional target.
fn split_scope(spec: &str) -> Result<(String, Option<&str>), ScopeError> {
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
            None => Err(ScopeError::TrailingText {
                rest: rest.to_string(),
                spec: spec.to_string(),
            }),
        };
    }
    Err(ScopeError::Unterminated {
        spec: spec.to_string(),
    })
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
    read_element_body_into(reader, name, &mut Vec::new(), &mut body)?;
    Ok(body)
}

/// As [`read_element_body`], into `body` (emptied first) and reading through `buf`: a loop over
/// many small elements (the cells of a worksheet) keeps both and allocates them once.
fn read_element_body_into<R: BufRead>(
    reader: &mut Reader<R>,
    name: &str,
    buf: &mut Vec<u8>,
    body: &mut Vec<Event<'static>>,
) -> Result<()> {
    body.clear();
    let mut depth = 0u32;
    loop {
        buf.clear();
        let event = reader.read_event_into(buf)?.into_owned();
        match &event {
            Event::Start(e) if e.name().as_ref() == name => depth += 1,
            Event::End(e) if e.name().as_ref() == name => {
                if depth == 0 {
                    return Ok(());
                }
                depth -= 1;
            }
            Event::Eof => {
                return Err(Error::malformed(format!(
                    "unexpected end of document inside <{name}>"
                )));
            }
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
            writer.write_event(Event::Start(start.borrow()))?;
            for event in content {
                writer.write_event(event.borrow())?;
            }
            writer.write_event(Event::End(start.to_end()))?;
        }
    }
    Ok(())
}

/// A destination that, once a write to it has failed, takes what comes after without writing it. The
/// archive writer finishes the archive when it is dropped and prints a line of its own to the standard
/// error if that fails too; the first error is already the one the caller is told of, so the rest
/// must be quiet.
pub(crate) struct Quiet<W> {
    inner: W,
    failed: bool,
}

impl<W> Quiet<W> {
    pub(crate) fn new(inner: W) -> Self {
        Self {
            inner,
            failed: false,
        }
    }
}

impl<W: std::io::Write> std::io::Write for Quiet<W> {
    fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
        if self.failed {
            return Ok(buf.len());
        }
        self.inner.write(buf).inspect_err(|_| self.failed = true)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        if self.failed {
            return Ok(());
        }
        self.inner.flush().inspect_err(|_| self.failed = true)
    }
}

impl<W: std::io::Seek> std::io::Seek for Quiet<W> {
    fn seek(&mut self, pos: std::io::SeekFrom) -> std::io::Result<u64> {
        if self.failed {
            return Ok(0);
        }
        self.inner.seek(pos).inspect_err(|_| self.failed = true)
    }
}

/// Modify a workbook file (supports multiple operations)
///
/// # Errors
///
/// When the input cannot be read or written, the format is not supported, an operation cannot be applied (it names a sheet the workbook does not have, or a deletion that cannot be made right), or a formula holds a 3D reference.
pub fn modify_workbook<P: AsRef<Path>>(
    input_path: P,
    output_path: P,
    modifications: &WorkbookModifications,
) -> Result<()> {
    let input = input_path.as_ref();
    let output = output_path.as_ref();

    // Creating the output truncates it before the input is read.
    if output.exists() && output.canonicalize()? == input.canonicalize()? {
        return Err(Error::OutputIsInput(output.to_path_buf()));
    }
    let input_file = File::open(input)?;
    // The format is the one of the content of the file, not of its extension.
    let format = Format::detect_reader(&input_file).map_err(|error| Error::NotAWorkbook {
        what: error.to_string(),
    })?;
    write_file(output, |output_file| match format {
        Format::Xlsx => modify_workbook_xlsx(input_file, output_file, modifications),
        Format::Ods => modify_workbook_ods(input_file, output_file, modifications),
        _ => Err(Error::UnsupportedFormat),
    })
}

impl Book {
    /// Modify the workbook and write the result to the file `output` as [`modify_workbook`] does
    /// (the same rules about what a failed run leaves behind). The pending operations of the book
    /// are not applied: they are for the planner, which hands the writer what each pass asks.
    ///
    /// # Errors
    ///
    /// As [`modify_workbook`].
    pub fn modify(&self, output: &Path, modifications: &WorkbookModifications) -> Result<()> {
        match (self.path(), self.bytes()) {
            (Some(path), _) => modify_workbook(path, output, modifications),
            (None, Some(data)) => write_file(output, |output_file| match self.format() {
                Format::Xlsx => modify_workbook_xlsx(
                    std::io::Cursor::new(data.as_ref()),
                    output_file,
                    modifications,
                ),
                Format::Ods => modify_workbook_ods(
                    std::io::Cursor::new(data.as_ref()),
                    output_file,
                    modifications,
                ),
                _ => Err(Error::UnsupportedFormat),
            }),
            (None, None) => unreachable!("a book is a file or bytes"),
        }
    }

    /// Modify the workbook and give the result in memory: no file is written, and nothing is read
    /// but the file the book names, if it is one.
    ///
    /// # Errors
    ///
    /// As [`modify_workbook_bytes`], and when the file cannot be read.
    pub fn modify_to_bytes(&self, modifications: &WorkbookModifications) -> Result<Vec<u8>> {
        match (self.path(), self.bytes()) {
            (Some(path), _) => {
                modify_workbook_bytes(self.format(), &std::fs::read(path)?, modifications)
            }
            (None, Some(data)) => modify_workbook_bytes(self.format(), data, modifications),
            (None, None) => unreachable!("a book is a file or bytes"),
        }
    }
}

/// Writes what `write_to` makes to the file `output`, deciding what a failed run leaves behind by
/// what was there before it. A file that was already there is never destroyed by it: it is replaced
/// only when the result is complete (written beside it and renamed over it), and a file that is not
/// a regular one (a device, a pipe) is written to and left alone. Only a file this command created
/// is removed when it fails.
fn write_file(output: &Path, write_to: impl FnOnce(File) -> Result<()>) -> Result<()> {
    match std::fs::metadata(output) {
        Ok(existing) if existing.is_file() => {
            let folder = output
                .parent()
                .filter(|p| !p.as_os_str().is_empty())
                .unwrap_or_else(|| Path::new("."));
            let temporary = tempfile::Builder::new()
                .prefix(".sheetrs-")
                .tempfile_in(folder)?;
            write_to(temporary.reopen()?)?;
            std::fs::set_permissions(temporary.path(), existing.permissions())?;
            temporary.persist(output).map_err(|e| e.error)?;
            Ok(())
        }
        Ok(_) => write_to(File::create(output)?),
        Err(_) => {
            let result = write_to(File::create(output)?);
            if result.is_err() {
                let _ = std::fs::remove_file(output);
            }
            result
        }
    }
}

/// Modify a workbook held in memory and give the result in memory: the same writers as
/// [`modify_workbook`], over the bytes of the archive.
///
/// # Errors
///
/// When the format is not one the writers support (`xlsx` and `ods`), the input cannot be read, an
/// operation cannot be applied, or a formula holds a 3D reference.
pub fn modify_workbook_bytes(
    format: Format,
    input: &[u8],
    modifications: &WorkbookModifications,
) -> Result<Vec<u8>> {
    let mut output = std::io::Cursor::new(Vec::new());
    match format {
        Format::Xlsx => {
            modify_workbook_xlsx(std::io::Cursor::new(input), &mut output, modifications)?;
        }
        Format::Ods => {
            modify_workbook_ods(std::io::Cursor::new(input), &mut output, modifications)?;
        }
        _ => return Err(Error::UnsupportedFormat),
    }
    Ok(output.into_inner())
}

/// The error value an XLSX writes for the error `code` of `LibreOffice`, which has no `Err:5xx`.
#[must_use]
pub fn xlsx_error(code: u16) -> &'static str {
    match code {
        503 => "#NUM!",
        507 | 520 | 32767 => "#N/A",
        521 => "#NULL!",
        524 => "#REF!",
        525 => "#NAME?",
        532 => "#DIV/0!",
        _ => "#VALUE!",
    }
}

/// The text an ODS shows for the error `code`: the name Excel gives it where `LibreOffice` shows one,
/// `Err:NNN` otherwise.
pub(crate) fn ods_error(code: u16) -> String {
    match code {
        503 => "#NUM!".into(),
        519 => "#VALUE!".into(),
        521 => "#NULL!".into(),
        524 => "#REF!".into(),
        525 => "#NAME?".into(),
        532 => "#DIV/0!".into(),
        32767 => "#N/A".into(),
        _ => format!("Err:{code}"),
    }
}

/// How a part that is rewritten is stored in the new archive: deflate at level 4, which leaves the
/// XML as small as the default level 6 (the same size to the permille on a 36 MB workbook) and
/// takes about a third less time to write; level 1 is faster still but makes the file 60% bigger.
/// The parts that are not rewritten are copied as they are, at their own level.
pub(crate) fn part_options() -> SimpleFileOptions {
    SimpleFileOptions::default().compression_level(Some(4))
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

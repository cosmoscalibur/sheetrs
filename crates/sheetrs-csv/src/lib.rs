//! Reader of delimited text files into the sheetrs model.
//!
//! A file is read as `LibreOffice` imports it with the text filter: the options below are the
//! ones of that filter. The file holds one sheet, which the caller names, because a name is not
//! part of the text.
//!
//! Dates, times, percentages, currencies and booleans are text, as the filter leaves them when
//! its option to detect special numbers is off; ISO 8601 dates and date-times are numbers in
//! either case.

mod date;
mod error;
mod field;
mod number;
mod write;

use std::borrow::Cow;
use std::io::Read;
use std::ops::ControlFlow;

use sheetrs_core::{CellPos, FormulaRef, MAX_COLS, MAX_ROWS, ValueRef, Workbook};

pub use error::Error;
pub use write::{WriteOptions, write};

/// What the filter makes of a field that is quoted.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum QuotedFields {
    /// A quoted field that looks like a number is a number.
    Detected,
    /// A quoted field is text.
    Text,
}

/// What the filter makes of a field that starts with `=`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaFields {
    /// It is text.
    Text,
    /// It is a formula, to be recalculated.
    Evaluated,
}

/// What the filter does with the spaces around a field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Spaces {
    /// They stay.
    Kept,
    /// They are dropped from fields that are not quoted.
    Trimmed,
}

/// The options of the import.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Options {
    /// The character that ends a field.
    pub separator: char,
    /// The character that quotes a field; `None` for no quoting.
    pub delimiter: Option<char>,
    pub quoted_fields: QuotedFields,
    pub formula_fields: FormulaFields,
    pub spaces: Spaces,
}

impl Options {
    /// Commas between fields, double quotes around them, and nothing else special.
    #[must_use]
    pub const fn comma_separated() -> Self {
        Self {
            separator: ',',
            delimiter: Some('"'),
            quoted_fields: QuotedFields::Detected,
            formula_fields: FormulaFields::Text,
            spaces: Spaces::Kept,
        }
    }
}

/// Reads a delimited text file of UTF-8 text into a workbook of one sheet called `sheet_name`.
///
/// Bytes that are not UTF-8 are read as the replacement character. The whole input is held in
/// memory while it is read.
///
/// # Errors
///
/// [`Error::Io`] when the input cannot be read, and [`Error::TooLarge`] when the text holds more
/// strings than the cell store can address.
pub fn read<R: Read>(
    mut source: R,
    sheet_name: &str,
    options: &Options,
) -> Result<Workbook, Error> {
    let mut bytes = Vec::new();
    source.read_to_end(&mut bytes)?;
    let lossy = String::from_utf8_lossy(&bytes);
    let text: Cow<'_, str> = if lossy.contains('\0') {
        Cow::Owned(lossy.replace('\0', ""))
    } else {
        lossy
    };
    let text = text.strip_prefix('\u{feff}').unwrap_or(&text);

    let mut workbook = Workbook::new();
    workbook.set_formula_grammar(sheetrs_core::FormulaGrammar::Native);
    let mut sheet = workbook.sheet_builder(sheet_name);
    let mut failure = None;
    field::fields(text, options, |row, col, field| {
        // Rows only grow, so nothing after the last row is read.
        if row >= MAX_ROWS {
            return ControlFlow::Break(());
        }
        if col >= MAX_COLS {
            return ControlFlow::Continue(());
        }
        let pos = CellPos { row, col };
        let outcome = match field::cell(&field, options) {
            field::Content::Empty => Ok(()),
            field::Content::Number(number) => {
                sheet.load_view(pos, &ValueRef::Number(number), None, false, None)
            }
            field::Content::Text(text) => {
                sheet.load_view(pos, &ValueRef::Text(text), None, false, None)
            }
            field::Content::Formula(text) => sheet.load_view(
                pos,
                &ValueRef::Empty,
                Some(FormulaRef::Text(text)),
                true,
                None,
            ),
        };
        match outcome {
            Ok(()) => ControlFlow::Continue(()),
            Err(error) => {
                failure = Some(error);
                ControlFlow::Break(())
            }
        }
    });
    if let Some(error) = failure {
        return Err(error.into());
    }
    sheet.finish();
    Ok(workbook)
}

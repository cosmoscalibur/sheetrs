//! Writer of a sheet as delimited text, as `LibreOffice` exports it with the text filter.

use std::io::Write;

use sheetrs_core::{SheetView, ValueRef};

use crate::Error;

/// The options of the export.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WriteOptions {
    /// The character that ends a field.
    pub separator: char,
    /// The character that quotes a field that needs it.
    pub delimiter: char,
    /// Whether every text cell is quoted, and not only those that need it.
    pub quote_all_text: bool,
    /// Whether the text starts with a byte order mark.
    pub byte_order_mark: bool,
}

impl Default for WriteOptions {
    /// A comma between fields and `"` around the fields that need it.
    fn default() -> Self {
        Self {
            separator: ',',
            delimiter: '"',
            quote_all_text: false,
            byte_order_mark: false,
        }
    }
}

/// Writes `sheet` to `destination`: one line per row from the first row to the last one that has
/// a cell, each with a field for every column up to the last one that has a cell. A cell shows as
/// General format shows it; text is quoted only when it holds the separator, the delimiter or a
/// line break.
///
/// # Errors
///
/// [`Error::Io`] when the destination cannot be written.
pub fn write<W: Write>(
    sheet: SheetView<'_>,
    options: &WriteOptions,
    mut destination: W,
) -> Result<(), Error> {
    let (mut last_row, mut last_col) = (None::<u32>, 0);
    for (pos, _) in sheet.cells_by_column() {
        last_row = Some(last_row.map_or(pos.row, |row| row.max(pos.row)));
        last_col = last_col.max(pos.col);
    }
    if options.byte_order_mark {
        destination.write_all("\u{feff}".as_bytes())?;
    }
    let Some(last_row) = last_row else {
        return Ok(());
    };
    let mut cells = sheet.cells().peekable();
    let mut line = String::new();
    for row in 0..=last_row {
        line.clear();
        for col in 0..=last_col {
            if col > 0 {
                line.push(options.separator);
            }
            if let Some((_, cell)) = cells.next_if(|(pos, _)| pos.row == row && pos.col == col) {
                push_field(&mut line, &cell.value, *options);
            }
        }
        line.push('\n');
        destination.write_all(line.as_bytes())?;
    }
    Ok(())
}

fn push_field(line: &mut String, value: &ValueRef<'_>, options: WriteOptions) {
    match value {
        ValueRef::Empty => {}
        ValueRef::Number(number) => line.push_str(&general(*number)),
        ValueRef::Bool(flag) => line.push_str(if *flag { "TRUE" } else { "FALSE" }),
        ValueRef::Error(kind) if options.quote_all_text => {
            push_quoted(line, kind.as_text(), options);
        }
        ValueRef::Error(kind) => line.push_str(kind.as_text()),
        ValueRef::Text(text) => {
            let needs_quotes = options.quote_all_text
                || text.chars().any(|character| {
                    character == options.separator
                        || character == options.delimiter
                        || character == '\n'
                        || character == '\r'
                });
            if needs_quotes {
                push_quoted(line, text, options);
            } else {
                line.push_str(text);
            }
        }
    }
}

/// Appends `text` between delimiters, with the delimiters inside it doubled.
fn push_quoted(line: &mut String, text: &str, options: WriteOptions) {
    line.push(options.delimiter);
    for character in text.chars() {
        if character == options.delimiter {
            line.push(options.delimiter);
        }
        line.push(character);
    }
    line.push(options.delimiter);
}

/// The number as General format shows it.
///
/// It has 15 significant digits. A number whose first digit is at power `e` of ten is written in
/// full from `1E-4` up to `1E+15`, and as a mantissa and an exponent from `1E+16` up. A number
/// below `1E-4` is written in full when it is at least `1E-9` and its decimals are no more than
/// 16, and otherwise as a mantissa and an exponent (`1E-10`, `1.5E+016`).
pub(crate) fn general(number: f64) -> String {
    let Some(found) = sheetrs_io::fifteen_digits(number) else {
        return "0".to_owned();
    };
    let length = i32::try_from(found.digits.len()).unwrap_or(i32::MAX);
    let plain = match found.exponent {
        16.. => false,
        -4..=15 => true,
        -9..=-5 => length - 1 - found.exponent <= 16,
        _ => false,
    };
    if plain {
        found.plain(number)
    } else {
        found.scientific(2)
    }
}

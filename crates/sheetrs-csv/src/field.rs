//! The fields of a delimited text, and what each of them holds.

use std::borrow::Cow;
use std::ops::ControlFlow;

use crate::{FormulaFields, Options, QuotedFields, Spaces, date, number};

/// What the fields read so far show about the quotes ahead.
#[derive(Default)]
struct Quotes {
    /// After a quoted field failed to close, the last quote that could close one: a field that
    /// opens beyond it cannot close either.
    horizon: Option<usize>,
}

/// One field of the text, with the quotes taken off.
pub(crate) struct Field<'a> {
    pub text: Cow<'a, str>,
    pub quoted: bool,
}

/// What a field puts in a cell.
pub(crate) enum Content<'a> {
    Empty,
    Number(f64),
    Text(&'a str),
    Formula(&'a str),
}

fn is_line_break(character: char) -> bool {
    character == '\r' || character == '\n'
}

/// The bytes of the line break at the start of `text`: a carriage return and a line feed, in
/// either order, are one.
fn line_break_len(text: &str) -> usize {
    let mut characters = text.chars();
    match (characters.next(), characters.next()) {
        (Some('\r'), Some('\n')) | (Some('\n'), Some('\r')) => 2,
        _ => 1,
    }
}

/// Calls `emit` with the row and column of each field of `text`, in order, until it breaks.
pub(crate) fn fields<'a>(
    text: &'a str,
    options: &Options,
    mut emit: impl FnMut(u32, u32, Field<'a>) -> ControlFlow<()>,
) {
    let (mut row, mut col) = (0_u32, 0_u32);
    let mut at = 0;
    let mut quotes = Quotes::default();
    while at < text.len() {
        let (field, end) = field(text, at, options, &mut quotes);
        if emit(row, col, field).is_break() {
            return;
        }
        match text.get(end..).and_then(|rest| rest.chars().next()) {
            None => return,
            Some(character) if character == options.separator => {
                at = end + character.len_utf8();
                col = col.saturating_add(1);
                if at == text.len() {
                    return;
                }
            }
            Some(_) => {
                at = end + line_break_len(text.get(end..).unwrap_or_default());
                row = row.saturating_add(1);
                col = 0;
            }
        }
    }
}

/// The field that starts at `at` and the position of what ends it: a separator, a line break or
/// the end of the text.
fn field<'a>(
    text: &'a str,
    at: usize,
    options: &Options,
    quotes: &mut Quotes,
) -> (Field<'a>, usize) {
    let rest = text.get(at..).unwrap_or_default();
    if let Some(delimiter) = options.delimiter {
        // A space that separates fields is not one of the spaces around a quoted field.
        let spaces = if options.separator == ' ' {
            0
        } else {
            rest.len() - rest.trim_start_matches(' ').len()
        };
        let after_spaces = at + spaces;
        let quoted_start = text
            .get(after_spaces..)
            .filter(|tail| tail.starts_with(delimiter))
            .map(|_| after_spaces + delimiter.len_utf8());
        if let Some((content, trailing, end)) =
            quoted_start.and_then(|start| quoted(text, start, delimiter, options.separator, quotes))
        {
            let mut field = content;
            if options.spaces == Spaces::Kept {
                field.push_str(trailing);
            }
            return (
                Field {
                    text: Cow::Owned(field),
                    quoted: true,
                },
                end,
            );
        }
    }
    let end = rest
        .find(|character: char| character == options.separator || is_line_break(character))
        .map_or(text.len(), |offset| at + offset);
    let raw = text.get(at..end).unwrap_or_default();
    let raw = if options.spaces == Spaces::Trimmed {
        raw.trim_matches(' ')
    } else {
        raw
    };
    (
        Field {
            text: Cow::Borrowed(raw),
            quoted: false,
        },
        end,
    )
}

/// The content of the quoted field whose first character after the opening quote is at `start`,
/// the spaces after its closing quote, and the position that ends the field.
///
/// A quote closes the field when only spaces separate it from a separator, a line break or the
/// end of the text; any other quote is part of the content, and two in a row are one. A field
/// with no such quote is not quoted.
fn quoted<'a>(
    text: &'a str,
    start: usize,
    delimiter: char,
    separator: char,
    quotes: &mut Quotes,
) -> Option<(String, &'a str, usize)> {
    if quotes.horizon.is_some_and(|horizon| start > horizon) {
        return None;
    }
    // The spaces after a quote and the position that follows them, if what follows ends a field.
    let ends_field = |after: usize| {
        let tail = text.get(after..)?;
        let spaces = if separator == ' ' {
            0
        } else {
            tail.len() - tail.trim_start_matches(' ').len()
        };
        let end = after + spaces;
        text.get(end..)?
            .chars()
            .next()
            .is_none_or(|character| character == separator || is_line_break(character))
            .then_some(end)
    };
    let mut at = start;
    let mut last_closer = None;
    let (close, after, end) = loop {
        let Some(offset) = text.get(at..).and_then(|tail| tail.find(delimiter)) else {
            quotes.horizon = Some(last_closer.unwrap_or(0));
            return None;
        };
        let found = at + offset;
        let after = found + delimiter.len_utf8();
        if text.get(after..)?.starts_with(delimiter) {
            let second = after + delimiter.len_utf8();
            if ends_field(second).is_some() {
                last_closer = Some(after);
            }
            at = second;
            continue;
        }
        if let Some(end) = ends_field(after) {
            break (found, after, end);
        }
        at = after;
    };
    let mut content = String::new();
    let mut rest = text.get(start..close)?;
    while let Some(offset) = rest.find(delimiter) {
        content.push_str(rest.get(..offset)?);
        content.push(delimiter);
        let skipped = offset + delimiter.len_utf8();
        rest = rest.get(skipped..)?;
        // Two quotes in a row are one.
        rest = rest.strip_prefix(delimiter).unwrap_or(rest);
    }
    content.push_str(rest);
    Some((normalize_line_breaks(&content), text.get(after..end)?, end))
}

/// Each line break of the text is a line feed.
fn normalize_line_breaks(text: &str) -> String {
    if !text.contains(is_line_break) {
        return text.to_owned();
    }
    let mut normalized = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(character) = rest.chars().next() {
        if is_line_break(character) {
            normalized.push('\n');
            rest = rest.get(line_break_len(rest)..).unwrap_or_default();
        } else {
            normalized.push(character);
            rest = rest.get(character.len_utf8()..).unwrap_or_default();
        }
    }
    normalized
}

/// What `field` puts in a cell.
pub(crate) fn cell<'a>(field: &'a Field<'_>, options: &Options) -> Content<'a> {
    let text: &str = &field.text;
    if text.is_empty() {
        return Content::Empty;
    }
    if options.formula_fields == FormulaFields::Evaluated && text.starts_with('=') && text.len() > 1
    {
        return Content::Formula(text);
    }
    let as_text = field.quoted && options.quoted_fields == QuotedFields::Text;
    if !as_text && let Some(value) = number::number(text).or_else(|| date::iso_date(text)) {
        return Content::Number(value);
    }
    Content::Text(text)
}

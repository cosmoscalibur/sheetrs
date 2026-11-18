//! XML helpers shared by the part parsers.

use std::borrow::Cow;
use std::io::BufRead;

use quick_xml::escape::resolve_xml_entity;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::{Reader, XmlVersion};

use crate::Error;

/// The value of the attribute whose local name is `local_name`, whatever its prefix.
pub(crate) fn attribute(
    start: &BytesStart<'_>,
    local_name: &str,
    part: &str,
) -> Result<Option<String>, Error> {
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| Error::malformed(part, error.to_string()))?;
        if attribute.key.local_name().as_ref() == local_name {
            let value = attribute
                .normalized_value(XmlVersion::Implicit1_0)
                .map_err(|source| Error::xml(part, source))?;
            return Ok(Some(value.into_owned()));
        }
    }
    Ok(None)
}

/// Decodes the `_xHHHH_` escapes of the `ST_Xstring` type, such as `_x000D_` for a carriage
/// return and `_x005F_` for an underscore. A high and a low surrogate escape in a row make one
/// character; a lone surrogate escape ends the text before it. Text that only resembles an
/// escape is kept.
pub(crate) fn unescape_xstring(text: &str) -> Cow<'_, str> {
    if !text.contains("_x") {
        return Cow::Borrowed(text);
    }
    let mut decoded = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(start) = rest.find("_x") {
        let (before, tail) = rest.split_at(start);
        decoded.push_str(before);
        match escape_at(tail) {
            Escape::Character(character, length) => {
                decoded.push(character);
                rest = tail.get(length..).unwrap_or_default();
            }
            Escape::LoneSurrogate => return Cow::Owned(decoded),
            Escape::Plain => {
                decoded.push_str("_x");
                rest = tail.get(2..).unwrap_or_default();
            }
        }
    }
    decoded.push_str(rest);
    Cow::Owned(decoded)
}

const ESCAPE_LEN: usize = "_x0000_".len();

enum Escape {
    /// A character and the bytes its escapes take.
    Character(char, usize),
    LoneSurrogate,
    Plain,
}

/// The code of the escape that `tail` starts with, if it starts with one.
fn escape_code(tail: &str) -> Option<u32> {
    let digits = tail.get(2..6)?;
    if tail.as_bytes().get(6) != Some(&b'_') || !digits.bytes().all(|byte| byte.is_ascii_hexdigit())
    {
        return None;
    }
    u32::from_str_radix(digits, 16).ok()
}

/// What the escape that `tail` may start with stands for.
fn escape_at(tail: &str) -> Escape {
    let Some(code) = escape_code(tail) else {
        return Escape::Plain;
    };
    let low_surrogate = |code: Option<u32>| code.filter(|code| (0xDC00..0xE000).contains(code));
    if (0xD800..0xDC00).contains(&code) {
        let low = tail
            .get(ESCAPE_LEN..)
            .and_then(escape_code)
            .and_then(|low| low_surrogate(Some(low)));
        return match low {
            Some(low) => {
                let combined = 0x1_0000 + ((code - 0xD800) << 10) + (low - 0xDC00);
                char::from_u32(combined).map_or(Escape::LoneSurrogate, |character| {
                    Escape::Character(character, 2 * ESCAPE_LEN)
                })
            }
            None => Escape::LoneSurrogate,
        };
    }
    match char::from_u32(code) {
        Some(character) => Escape::Character(character, ESCAPE_LEN),
        None => Escape::LoneSurrogate,
    }
}

/// The text of the element just opened, read up to its end tag.
pub(crate) fn element_text<R: BufRead>(
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
    part: &str,
) -> Result<String, Error> {
    let mut text = String::new();
    let mut depth = 0_usize;
    loop {
        buf.clear();
        match reader
            .read_event_into(buf)
            .map_err(|source| Error::xml(part, source))?
        {
            Event::Text(content) => text.push_str(&content.xml10_content()),
            Event::CData(content) => text.push_str(&content.xml10_content()),
            Event::GeneralRef(reference) => push_reference(&mut text, &reference, part)?,
            Event::Start(_) => depth += 1,
            Event::End(_) => {
                if depth == 0 {
                    return Ok(text);
                }
                depth -= 1;
            }
            Event::Eof => return Err(Error::malformed(part, "unexpected end of file")),
            _ => {}
        }
    }
}

/// Skips the element just opened, including everything inside it.
pub(crate) fn skip_element<R: BufRead>(
    reader: &mut Reader<R>,
    start: &BytesStart<'_>,
    part: &str,
) -> Result<(), Error> {
    let mut scratch = Vec::new();
    reader
        .read_to_end_into(start.name(), &mut scratch)
        .map_err(|source| Error::xml(part, source))?;
    Ok(())
}

fn push_reference(text: &mut String, reference: &BytesRef<'_>, part: &str) -> Result<(), Error> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|source| Error::xml(part, source))?
    {
        text.push(character);
        return Ok(());
    }
    let name = reference.xml10_content();
    let resolved = resolve_xml_entity(&name)
        .ok_or_else(|| Error::malformed(part, format!("unknown entity &{name};")))?;
    text.push_str(resolved);
    Ok(())
}

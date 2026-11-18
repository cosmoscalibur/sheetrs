//! The text of a string as an XLSX file writes it.

use std::fmt::Write as _;

/// Appends `text` to `out` escaped for the content or an attribute of an XML element, with the
/// characters XLSX cannot hold written as `_xHHHH_`, and a literal `_xHHHH_` with its underscore
/// escaped.
pub(crate) fn push_text(out: &mut String, text: &str) {
    for (at, character) in text.char_indices() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            '\n' => out.push_str("&#10;"),
            '\r' => out.push_str("&#13;"),
            '\t' => out.push_str("&#9;"),
            '_' if looks_like_escape(text.get(at..).unwrap_or_default()) => {
                out.push_str("_x005F_");
            }
            character if is_xml_character(character) => out.push(character),
            character => {
                let mut units = [0_u16; 2];
                for unit in character.encode_utf16(&mut units) {
                    let _ = write!(out, "_x{unit:04X}_");
                }
            }
        }
    }
}

/// Whether `text` starts with `_xHHHH_`, which a reader would take for an escape.
fn looks_like_escape(text: &str) -> bool {
    let bytes = text.as_bytes();
    bytes.len() >= 7
        && bytes.first() == Some(&b'_')
        && bytes.get(1) == Some(&b'x')
        && bytes
            .get(2..6)
            .is_some_and(|hex| hex.iter().all(u8::is_ascii_hexdigit))
        && bytes.get(6) == Some(&b'_')
}

/// Whether XML 1.0 allows the character, apart from the whitespace written as references.
fn is_xml_character(character: char) -> bool {
    matches!(character, '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
}

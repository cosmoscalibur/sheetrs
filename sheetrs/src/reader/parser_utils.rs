//! Common parsing utilities shared between XLSX and ODS parsers

use std::borrow::Cow;

use anyhow::Result;
use quick_xml::events::Event;
use quick_xml::events::attributes::Attribute;
use quick_xml::{Reader, XmlVersion};

/// Parse a cell reference like "A1" into (row, col) as 0-based indices
///
/// # Arguments
///
/// * `cell_ref` - The cell reference string (e.g., "A1", "Z99")
///
/// # Returns
///
/// * `Option<(u32, u32)>` - The (row, col) indices if parsing is successful.
pub fn parse_cell_ref(cell_ref: &str) -> Option<(u32, u32)> {
    let row_str: String = cell_ref.chars().filter(char::is_ascii_digit).collect();
    let row = row_str.parse::<u32>().ok()?;
    let col = column_letters_to_index(cell_ref).unwrap_or(0);

    // Convert to 0-based
    Some((row.saturating_sub(1), col))
}

/// Parse a cell range like "A1:B2" into (start_row, start_col, end_row, end_col)
///
/// # Arguments
///
/// * `range` - The range string (e.g., "A1:B2")
///
/// # Returns
///
/// * `Option<(u32, u32, u32, u32)>` - The range coordinates if parsing is successful.
pub fn parse_cell_range(range: &str) -> Option<(u32, u32, u32, u32)> {
    let parts: Vec<&str> = range.split(':').collect();
    if parts.len() != 2 {
        return None;
    }

    let (start_row, start_col) = parse_cell_ref(parts[0])?;
    let (end_row, end_col) = parse_cell_ref(parts[1])?;

    Some((start_row, start_col, end_row, end_col))
}

/// Converts the column letters in `reference` to a 0-based column index
/// (`A` → 0, `AA` → 26), case-insensitively.
///
/// Non-letter characters are skipped, so a full reference such as `$AA$1`
/// yields its column. Returns `None` when `reference` has no letters.
pub fn column_letters_to_index(reference: &str) -> Option<u32> {
    reference
        .bytes()
        .filter(u8::is_ascii_alphabetic)
        .fold(0u32, |acc, b| {
            acc * 26 + u32::from(b.to_ascii_uppercase() - b'A') + 1
        })
        .checked_sub(1)
}

/// Converts a 0-based column index to column letters (`0` → `A`, `26` → `AA`).
pub fn column_index_to_letters(mut index: u32) -> String {
    let mut letters = Vec::new();
    loop {
        letters.push(b'A' + (index % 26) as u8);
        if index < 26 {
            break;
        }
        index = index / 26 - 1;
    }
    letters.iter().rev().map(|&b| char::from(b)).collect()
}

/// Returns the value of an XML attribute with entities resolved.
///
/// Applies XML 1.0 attribute-value normalization: literal tabs and line breaks
/// become spaces, while character references such as `&#10;` are kept.
pub fn attr_value<'a>(attr: &Attribute<'a>) -> quick_xml::Result<Cow<'a, str>> {
    attr.normalized_value(XmlVersion::Implicit1_0)
}

/// Read text content from an XML node
pub fn read_text_node<R: std::io::BufRead>(reader: &mut Reader<R>) -> Result<String> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Text(e) => text.push_str(&e),
            Event::CData(e) => text.push_str(&e),
            Event::End(_) => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

/// Extract basename from a file path
/// Handles both absolute and relative paths, and strips file:// prefix
pub fn extract_basename(path: &str) -> String {
    // Strip file:// prefix if present
    let clean_path = path.trim_start_matches("file://");

    // Extract basename using Path
    std::path::Path::new(clean_path)
        .file_name()
        .and_then(|s| s.to_str())
        .unwrap_or(clean_path)
        .to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("B2"), Some((1, 1)));
        assert_eq!(parse_cell_ref("Z26"), Some((25, 25)));
        assert_eq!(parse_cell_ref("AA1"), Some((0, 26)));
        assert_eq!(parse_cell_ref("AB10"), Some((9, 27)));
    }

    #[test]
    fn test_column_letters_to_index() {
        assert_eq!(column_letters_to_index("A"), Some(0));
        assert_eq!(column_letters_to_index("Z"), Some(25));
        assert_eq!(column_letters_to_index("AA"), Some(26));
        assert_eq!(column_letters_to_index("XFD"), Some(16383));
        assert_eq!(column_letters_to_index("xfd"), Some(16383));
        assert_eq!(column_letters_to_index("$AA$1"), Some(26));
        assert_eq!(column_letters_to_index(""), None);
        assert_eq!(column_letters_to_index("12"), None);
    }

    #[test]
    fn test_column_index_to_letters() {
        assert_eq!(column_index_to_letters(0), "A");
        assert_eq!(column_index_to_letters(25), "Z");
        assert_eq!(column_index_to_letters(26), "AA");
        assert_eq!(column_index_to_letters(16383), "XFD");
    }

    #[test]
    fn test_column_conversion_round_trip() {
        for index in 0..=16383 {
            let letters = column_index_to_letters(index);
            assert_eq!(column_letters_to_index(&letters), Some(index), "{letters}");
        }
    }

    fn first_attr_value(xml: &str) -> quick_xml::Result<String> {
        let mut reader = Reader::from_str(xml);
        let Event::Empty(e) = reader.read_event()? else {
            panic!("expected an empty element");
        };
        let attr = e.attributes().next().expect("attribute")?;
        Ok(attr_value(&attr)?.into_owned())
    }

    #[test]
    fn test_attr_value_resolves_entities() {
        assert_eq!(
            first_attr_value(r#"<a v="R&amp;D &lt;1&gt;"/>"#).unwrap(),
            "R&D <1>"
        );
    }

    #[test]
    fn test_attr_value_keeps_char_ref_line_breaks() {
        assert_eq!(first_attr_value(r#"<a v="x&#10;y"/>"#).unwrap(), "x\ny");
    }

    #[test]
    fn test_attr_value_normalizes_literal_whitespace() {
        assert_eq!(first_attr_value("<a v=\"x\ny\tz\"/>").unwrap(), "x y z");
    }

    #[test]
    fn test_read_text_node_text_and_cdata() {
        let mut reader = Reader::from_str("<t>a b<![CDATA[<c>]]></t>");
        reader.read_event().unwrap();
        assert_eq!(read_text_node(&mut reader).unwrap(), "a b<c>");
    }

    #[test]
    fn test_read_text_node_rejects_invalid_utf8() {
        let xml: &[u8] = b"<t>a\xFFb</t>";
        let mut reader = Reader::from_reader(xml);
        let mut buf = Vec::new();
        reader.read_event_into(&mut buf).unwrap();
        assert!(read_text_node(&mut reader).is_err());
    }

    #[test]
    fn test_parse_cell_range() {
        assert_eq!(parse_cell_range("A1:B2"), Some((0, 0, 1, 1)));
        assert_eq!(parse_cell_range("C3:D4"), Some((2, 2, 3, 3)));
        assert_eq!(parse_cell_range("A1:Z26"), Some((0, 0, 25, 25)));
    }
}

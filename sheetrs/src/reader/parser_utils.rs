//! Common parsing utilities shared between XLSX and ODS parsers

use anyhow::Result;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Parse a cell reference like "A1" into (row, col) as 0-based indices
pub fn parse_cell_ref(cell_ref: &str) -> Option<(u32, u32)> {
    let mut col = 0u32;
    let mut row_str = String::new();

    for ch in cell_ref.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else if ch.is_ascii_digit() {
            row_str.push(ch);
        }
    }

    if row_str.is_empty() {
        return None;
    }

    let row = row_str.parse::<u32>().ok()?;

    // Convert to 0-based
    Some((row.saturating_sub(1), col.saturating_sub(1)))
}

/// Parse a cell range like "A1:B2" into (start_row, start_col, end_row, end_col)
pub fn parse_cell_range(range: &str) -> Option<(u32, u32, u32, u32)> {
    let parts: Vec<&str> = range.split(':').collect();
    if parts.len() != 2 {
        return None;
    }

    let (start_row, start_col) = parse_cell_ref(parts[0])?;
    let (end_row, end_col) = parse_cell_ref(parts[1])?;

    Some((start_row, start_col, end_row, end_col))
}

/// Read text content from an XML node
pub fn read_text_node<R: std::io::BufRead>(reader: &mut Reader<R>) -> Result<String> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Text(e) => text.push_str(e.unescape()?.as_ref()),
            Event::CData(e) => text.push_str(&String::from_utf8_lossy(e.as_ref())),
            Event::End(_) => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
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
    fn test_parse_cell_range() {
        assert_eq!(parse_cell_range("A1:B2"), Some((0, 0, 1, 1)));
        assert_eq!(parse_cell_range("C3:D4"), Some((2, 2, 3, 3)));
        assert_eq!(parse_cell_range("A1:Z26"), Some((0, 0, 25, 25)));
    }
}

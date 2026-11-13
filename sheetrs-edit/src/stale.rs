//! Whether the results a workbook caches are known to be out of date.
//!
//! A workbook that was saved after a change that leaves results stale (rows deleted, a name or a
//! sheet removed) says so: an XLSX carries `fullCalcOnLoad` in `calcPr`, and an ODS, which has no
//! such mark, has no cached result at all in its formula cells (the application computes them when
//! it opens the file). A conversion of formulas to values reads those results, so before it runs
//! on such a workbook the results have to be calculated again.

use crate::error::Result;
use quick_xml::Reader;
use quick_xml::events::Event;
use sheetrs_io::Book;
use sheetrs_io::reader::Format;
use std::io::{BufReader, Cursor, Read, Seek};

/// Whether the cached results of `source` are known to be stale.
///
/// # Errors
///
/// When the file cannot be read or is not a workbook of a format the editor writes.
pub(crate) fn input_is_stale(source: &Book) -> Result<bool> {
    let format = source.format();
    match (source.path(), source.bytes()) {
        (Some(path), _) => scan(std::fs::File::open(path)?, format),
        (None, Some(data)) => scan(Cursor::new(data.as_ref()), format),
        (None, None) => unreachable!("a book is a file or bytes"),
    }
}

fn scan<R: Read + Seek>(reader: R, format: Format) -> Result<bool> {
    let mut archive = zip::ZipArchive::new(reader)?;
    match format {
        Format::Xlsx => {
            let mut xml = String::new();
            archive
                .by_name("xl/workbook.xml")?
                .read_to_string(&mut xml)?;
            Ok(xlsx_marks_recalculation(&xml))
        }
        Format::Ods => {
            // Almost every file has a result in each formula cell: a scan of the bytes tells it
            // without building an XML event per element; the reader below decides the rest.
            if !ods_may_have_a_formula_without_result(archive.by_name("content.xml")?)? {
                return Ok(false);
            }
            let entry = archive.by_name("content.xml")?;
            ods_has_a_formula_without_result(BufReader::new(entry))
        }
        _ => Err(sheetrs_io::write::Error::UnsupportedFormat.into()),
    }
}

/// Whether `workbook.xml` asks to recalculate when the workbook is opened.
fn xlsx_marks_recalculation(xml: &str) -> bool {
    let mut reader = Reader::from_str(xml);
    while let Ok(event) = reader.read_event() {
        match event {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "calcPr" => {
                return e.attributes().flatten().any(|a| {
                    a.key.as_ref() == "fullCalcOnLoad" && matches!(a.value.as_ref(), "1" | "true")
                });
            }
            Event::Eof => break,
            _ => {}
        }
    }
    false
}

const CELL_START: &[u8] = b"<table:table-cell";

/// Where the start tag that begins at `from` ends: just past its `>`, which an attribute value
/// may hold between its quotes. `None` when `bytes` ends before it.
fn tag_end(bytes: &[u8], from: usize) -> Option<usize> {
    let mut quote = None;
    for (i, &byte) in bytes.iter().enumerate().skip(from) {
        match (quote, byte) {
            (Some(q), b) if b == q => quote = None,
            (None, b'"' | b'\'') => quote = Some(byte),
            (None, b'>') => return Some(i + 1),
            (Some(_) | None, _) => {}
        }
    }
    None
}

/// Whether the attributes of a cell (`attributes`, the bytes after its name up to the `>`) name a
/// formula and no value type.
fn formula_without_value_type_in(attributes: &[u8]) -> bool {
    let (mut formula, mut value_type) = (false, false);
    let mut i = 0;
    while i < attributes.len() {
        while i < attributes.len() && (attributes[i].is_ascii_whitespace() || attributes[i] == b'/')
        {
            i += 1;
        }
        let name_start = i;
        while i < attributes.len() && attributes[i] != b'=' && !attributes[i].is_ascii_whitespace()
        {
            i += 1;
        }
        let name = &attributes[name_start..i];
        while i < attributes.len() && (attributes[i] == b'=' || attributes[i].is_ascii_whitespace())
        {
            i += 1;
        }
        let Some(&quote) = attributes.get(i) else {
            break;
        };
        // The value runs to the matching quote.
        i += 1;
        while i < attributes.len() && attributes[i] != quote {
            i += 1;
        }
        i += 1;
        formula |= name == b"table:formula";
        value_type |= name == b"office:value-type";
    }
    formula && !value_type
}

/// Whether `content.xml` may hold a formula cell with no cached result, looking only at the bytes of
/// the start tags of its cells. `false` is certain; `true` is for the reader of
/// [`ods_has_a_formula_without_result`] to decide (a cell with no value type may still hold a
/// paragraph with its result, and a file that is not well formed is its error to report).
fn ods_may_have_a_formula_without_result<R: Read>(mut input: R) -> Result<bool> {
    let finder = aho_corasick::AhoCorasick::new([CELL_START]).expect("a literal pattern");
    let mut window: Vec<u8> = Vec::new();
    let mut chunk = vec![0u8; 1 << 20];
    loop {
        let read = input.read(&mut chunk)?;
        window.extend_from_slice(&chunk[..read]);
        let mut at = 0;
        loop {
            let Some(found) = finder.find(&window[at..]) else {
                // The end of the window may hold the start of a name.
                at = window.len().saturating_sub(CELL_START.len());
                break;
            };
            let start = at + found.start();
            let name_end = start + CELL_START.len();
            let Some(end) = tag_end(&window, name_end) else {
                at = start;
                break;
            };
            // `<table:table-cell-range-source` is not a cell.
            if matches!(window[name_end], b' ' | b'\t' | b'\n' | b'\r' | b'>' | b'/')
                && formula_without_value_type_in(&window[name_end..end - 1])
            {
                return Ok(true);
            }
            at = end;
        }
        window.drain(..at);
        if read == 0 {
            // Anything left is a tag that never ends.
            return Ok(window.windows(CELL_START.len()).any(|w| w == CELL_START));
        }
    }
}

/// Whether a cell of `content.xml` has a formula and no cached result: no value type, and no text
/// paragraph either (`LibreOffice` writes an empty text result as a paragraph with no value type).
fn ods_has_a_formula_without_result<R: std::io::BufRead>(input: R) -> Result<bool> {
    let mut reader = Reader::from_reader(input);
    let mut buffer = Vec::new();
    // Inside a formula cell that has no value type, waiting for a paragraph or the end of the cell.
    let mut waiting = false;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(e) if e.name().as_ref() == "table:table-cell" => {
                waiting = formula_without_value_type(&e)?;
            }
            Event::Empty(e) if e.name().as_ref() == "table:table-cell" => {
                if formula_without_value_type(&e)? {
                    return Ok(true);
                }
            }
            Event::Start(e) | Event::Empty(e) if waiting && e.name().as_ref() == "text:p" => {
                waiting = false;
            }
            Event::End(e) if e.name().as_ref() == "table:table-cell" => {
                if waiting {
                    return Ok(true);
                }
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn formula_without_value_type(cell: &quick_xml::events::BytesStart<'_>) -> Result<bool> {
    Ok(cell.try_get_attribute("table:formula")?.is_some()
        && cell.try_get_attribute("office:value-type")?.is_none())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn an_xlsx_that_asks_to_recalculate_when_it_is_opened_is_stale() {
        let marked = r#"<workbook><sheets/><calcPr calcId="1" fullCalcOnLoad="1"/></workbook>"#;
        let truthy = r#"<workbook><calcPr fullCalcOnLoad="true"></calcPr></workbook>"#;
        let plain = r#"<workbook><calcPr calcId="1"/></workbook>"#;
        assert!(xlsx_marks_recalculation(marked));
        assert!(xlsx_marks_recalculation(truthy));
        assert!(!xlsx_marks_recalculation(plain));
        assert!(!xlsx_marks_recalculation("<workbook/>"));
    }

    fn ods(cells: &str) -> bool {
        let xml = format!("<table:table><table:table-row>{cells}</table:table-row></table:table>");
        ods_has_a_formula_without_result(xml.as_bytes()).unwrap()
    }

    #[test]
    fn an_ods_formula_with_no_cached_result_is_stale() {
        assert!(ods(r#"<table:table-cell table:formula="of:=1+1"/>"#));
        assert!(ods(
            r#"<table:table-cell table:formula="of:=1+1"></table:table-cell>"#
        ));
    }

    /// The scan of the bytes never says `false` where the reader says `true`, whatever the size of
    /// the pieces the file arrives in.
    #[test]
    fn the_scan_of_the_bytes_agrees_with_the_reader() {
        struct Pieces<'a>(&'a [u8], usize);
        impl Read for Pieces<'_> {
            fn read(&mut self, out: &mut [u8]) -> std::io::Result<usize> {
                let n = self.1.min(self.0.len()).min(out.len());
                out[..n].copy_from_slice(&self.0[..n]);
                self.0 = &self.0[n..];
                Ok(n)
            }
        }
        let cases = [
            (r#"<table:table-cell table:formula="of:=1+1"/>"#, true),
            (
                r#"<table:table-cell table:formula="of:=1+1"></table:table-cell>"#,
                true,
            ),
            (
                r#"<table:table-cell table:formula="of:=A&gt;1" office:value-type="float" office:value="2"/>"#,
                false,
            ),
            // A value type named inside a quoted value is not an attribute.
            (
                r#"<table:table-cell table:formula='of:=" office:value-type="'/>"#,
                true,
            ),
            (
                r#"<table:table-cell-range-source table:formula="x"/>"#,
                false,
            ),
            (
                r#"<table:table-cell office:value-type="string"><text:p>table:formula=</text:p></table:table-cell>"#,
                false,
            ),
        ];
        for (cells, expected) in cases {
            let xml = format!("<table:table-row>{cells}</table:table-row>");
            for piece in [1, 2, 7, 1 << 20] {
                let got =
                    ods_may_have_a_formula_without_result(Pieces(xml.as_bytes(), piece)).unwrap();
                assert_eq!(got, expected, "{cells} in pieces of {piece}");
            }
        }
    }

    #[test]
    fn an_ods_formula_with_a_result_is_not() {
        assert!(!ods(
            r#"<table:table-cell table:formula="of:=1+1" office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell>"#
        ));
        // An empty text result is a paragraph with no value type.
        assert!(!ods(
            r#"<table:table-cell table:formula="of:=&quot;&quot;"><text:p/></table:table-cell>"#
        ));
        assert!(!ods(
            r#"<table:table-cell office:value-type="float" office:value="2"/>"#
        ));
    }
}

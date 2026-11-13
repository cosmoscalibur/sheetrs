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

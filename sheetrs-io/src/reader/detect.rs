//! The format of a workbook, read from its content.
//!
//! A workbook says what it is: an `OpenDocument` package carries a `mimetype` entry and an Office
//! Open XML package carries `[Content_Types].xml`, whose entry for `/xl/workbook.xml` gives the
//! content type of the main part. The name of the file is not consulted: a file called `.xlsx` that
//! holds an ODS is an ODS.

use super::Format;
use crate::error::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::{Read, Seek};
use zip::ZipArchive;

const ODS_MIME: &str = "application/vnd.oasis.opendocument.spreadsheet";
const XLSX_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const XLSM_MAIN: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";

/// The format of the workbook whose archive is `archive`.
///
/// # Errors
///
/// [`Error::Unsupported`] naming what the package is when it is not an XLSX, XLSM or ODS workbook
/// (another `OpenDocument` or Office Open XML document, an Excel template or binary workbook, a zip
/// that is neither), and the error of the archive when it cannot be read.
pub(crate) fn detect_archive<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<Format> {
    // OpenDocument: the `mimetype` entry names the type of the package.
    if let Some(mime) = entry_text(archive, "mimetype")? {
        let mime = mime.trim();
        if mime == ODS_MIME {
            return Ok(Format::Ods);
        }
        if mime.starts_with("application/vnd.oasis.opendocument.") {
            return Err(Error::unsupported(format!(
                "an OpenDocument file of type '{mime}', not an ODS spreadsheet"
            )));
        }
    }
    // Office Open XML: the content type of the workbook part.
    if let Some(types) = entry_text(archive, "[Content_Types].xml")? {
        return from_content_types(&types, archive.by_name("xl/workbook.xml").is_ok());
    }
    // An OpenDocument package written without `mimetype`.
    if archive.by_name("content.xml").is_ok() {
        return Ok(Format::Ods);
    }
    Err(Error::unsupported(
        "a zip archive that is neither an OpenDocument package (no `mimetype` entry) nor an Office \
         Open XML package (no `[Content_Types].xml` entry): not an XLSX, XLSM or ODS workbook",
    ))
}

/// The format of the workbook in `reader`.
///
/// # Errors
///
/// As [`detect_archive`]; [`Error::Unsupported`] also for input that is empty or not a zip archive.
pub(crate) fn detect_reader<R: Read + Seek>(mut reader: R) -> Result<Format> {
    let length = reader.seek(std::io::SeekFrom::End(0))?;
    reader.rewind()?;
    if length == 0 {
        return Err(Error::unsupported(
            "empty input: not an XLSX, XLSM or ODS workbook",
        ));
    }
    let mut archive = ZipArchive::new(reader).map_err(|error| {
        Error::unsupported(format!(
            "not a zip archive, or a truncated one ({error}): not an XLSX, XLSM or ODS workbook"
        ))
    })?;
    detect_archive(&mut archive)
}

/// `has_workbook_part` is whether the archive holds `xl/workbook.xml`: a package that gives no
/// content type for it (a producer that leaves the `Override` out) is read as an XLSX, since the
/// part is there.
fn from_content_types(types: &str, has_workbook_part: bool) -> Result<Format> {
    let mut reader = Reader::from_str(types);
    let (mut workbook, mut other) = (None, None);
    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == "Override" => {
                let part = attribute(&e, "PartName")?;
                let content = attribute(&e, "ContentType")?;
                if let (Some(part), Some(content)) = (part, content) {
                    match part.as_str() {
                        "/xl/workbook.xml" => workbook = Some(content),
                        "/word/document.xml" => other = Some("a Word document"),
                        "/ppt/presentation.xml" => other = Some("a PowerPoint presentation"),
                        _ => {}
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    match workbook.as_deref() {
        Some(XLSX_MAIN) => Ok(Format::Xlsx),
        Some(XLSM_MAIN) => Ok(Format::Xlsm),
        Some("application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml") => {
            Err(Error::unsupported(
                "an Excel template (XLTX), not an XLSX, XLSM or ODS workbook",
            ))
        }
        Some("application/vnd.ms-excel.template.macroEnabled.main+xml") => Err(Error::unsupported(
            "an Excel macro-enabled template (XLTM), not an XLSX, XLSM or ODS workbook",
        )),
        Some("application/vnd.ms-excel.sheet.binary.macroEnabled.main+xml") => {
            Err(Error::unsupported(
                "an Excel binary workbook (XLSB), not an XLSX, XLSM or ODS workbook",
            ))
        }
        Some("application/vnd.ms-excel.addin.macroEnabled.main+xml") => Err(Error::unsupported(
            "an Excel add-in (XLAM), not an XLSX, XLSM or ODS workbook",
        )),
        Some(content) => Err(Error::unsupported(format!(
            "an Office Open XML workbook of content type '{content}', not an XLSX, XLSM or ODS workbook"
        ))),
        None if has_workbook_part => Ok(Format::Xlsx),
        None => Err(Error::unsupported(format!(
            "{}: an Office Open XML package without `/xl/workbook.xml`, not an XLSX, XLSM or ODS workbook",
            other.unwrap_or("an Office Open XML package that is not a spreadsheet")
        ))),
    }
}

fn attribute(e: &quick_xml::events::BytesStart, key: &str) -> Result<Option<String>> {
    Ok(match e.try_get_attribute(key)? {
        Some(a) => Some(super::parser_utils::attr_value(&a)?.into_owned()),
        None => None,
    })
}

/// The text of the entry `name`, `None` when the archive has none (a read of at most 1 MiB: the
/// entries asked for are tiny).
fn entry_text<R: Read + Seek>(archive: &mut ZipArchive<R>, name: &str) -> Result<Option<String>> {
    let Ok(entry) = archive.by_name(name) else {
        return Ok(None);
    };
    let mut text = String::new();
    entry.take(1 << 20).read_to_string(&mut text)?;
    Ok(Some(text))
}

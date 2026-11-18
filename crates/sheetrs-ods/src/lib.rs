//! Streaming reader of ODS workbooks into the sheetrs model.
//!
//! The cells are in `content.xml`, which is read once without loading the whole part into
//! memory. Cell contents are read as `LibreOffice` loads them. A formula keeps the text of its
//! `table:formula` attribute, with its namespace prefix: the grammar of the formula is the one
//! the prefix names.

mod content;
mod error;
mod tokens;
mod value;

use std::io::{BufReader, Read, Seek};

use sheetrs_core::Workbook;
use zip::ZipArchive;
use zip::result::ZipError;

pub use error::Error;

const CONTENT_PART: &str = "content.xml";
const BUFFER_BYTES: usize = 64 * 1024;

/// Reads a workbook from an ODS package.
///
/// # Errors
///
/// [`Error::Zip`] when the input is not a ZIP package, [`Error::MissingPart`] when the package
/// has no `content.xml`, and [`Error::Xml`] or [`Error::Malformed`] when it cannot be read.
pub fn read<R: Read + Seek>(source: R) -> Result<Workbook, Error> {
    let mut archive = ZipArchive::new(source)?;
    let part = match archive.by_name(CONTENT_PART) {
        Ok(part) => part,
        Err(ZipError::FileNotFound) => return Err(Error::MissingPart(CONTENT_PART.to_owned())),
        Err(error) => return Err(error.into()),
    };
    let mut workbook = Workbook::new();
    content::parse(BufReader::with_capacity(BUFFER_BYTES, part), &mut workbook)?;
    Ok(workbook)
}

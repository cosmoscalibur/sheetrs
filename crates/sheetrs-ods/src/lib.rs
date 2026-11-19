//! Streaming reader of ODS workbooks into the sheetrs model.
//!
//! The cells are in `content.xml`, which is read once without loading the whole part into
//! memory. Cell contents are read as `LibreOffice` loads them. A formula keeps the text of its
//! `table:formula` attribute, with its namespace prefix: the grammar of the formula is the one
//! the prefix names.

mod content;
mod error;
mod macros;
mod tokens;
mod value;
mod write;

use std::io::{BufReader, Cursor, Read, Seek};

use sheetrs_core::{Frame, OpaquePackage, OpaquePart, PackageFormat, Pristine, Workbook};
use zip::ZipArchive;
use zip::result::ZipError;
use zip::write::ZipWriter;

pub use error::Error;
pub use write::write;

const CONTENT_PART: &str = "content.xml";
/// The most bytes of package entries that a workbook carries to write them back.
const MAX_OPAQUE_BYTES: u64 = 1 << 29;
const BUFFER_BYTES: usize = 64 * 1024;

/// Reads a workbook from an ODS package.
///
/// # Errors
///
/// [`Error::Zip`] when the input is not a ZIP package, [`Error::MissingPart`] when the package
/// has no `content.xml`, and [`Error::Xml`] or [`Error::Malformed`] when it cannot be read.
pub fn read<R: Read + Seek>(source: R) -> Result<Workbook, Error> {
    let mut archive = ZipArchive::new(source)?;
    let mut workbook = Workbook::new();
    workbook.set_formula_grammar(sheetrs_core::FormulaGrammar::OpenFormula);
    {
        let part = match archive.by_name(CONTENT_PART) {
            Ok(part) => part,
            Err(ZipError::FileNotFound) => {
                return Err(Error::MissingPart(CONTENT_PART.to_owned()));
            }
            Err(error) => return Err(error.into()),
        };
        content::parse(BufReader::with_capacity(BUFFER_BYTES, part), &mut workbook)?;
    }
    macros::list(&mut archive, &mut workbook)?;
    // Entries too large to keep are not carried: the writer then builds the package.
    if let Some((parts, entry)) = unread_parts(&mut archive)? {
        workbook.set_opaque(OpaquePackage {
            format: PackageFormat::Ods,
            parts,
            frame: Frame::default(),
            date_system: workbook.date_system(),
            defined_names: Vec::new(),
            pristine: Some(Pristine {
                entry,
                fingerprint: workbook.fingerprint(),
            }),
            shared_strings: None,
        });
    }
    Ok(workbook)
}

/// The entries of a package that the model does not interpret, and `content.xml` compressed.
type Unread = (Vec<OpaquePart>, Vec<u8>);

/// The entries of the package other than `content.xml`, as they are, and `content.xml` as it is,
/// compressed and alone in a ZIP archive; `None` when the entries hold more than
/// [`MAX_OPAQUE_BYTES`].
fn unread_parts<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<Option<Unread>, Error> {
    let mut parts = Vec::new();
    let mut total: u64 = 0;
    let mut content = ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let file = archive.by_index_raw(index)?;
        let name = file.name()?.into_owned();
        if file.is_dir() {
            continue;
        }
        if name == CONTENT_PART {
            content.raw_copy_file(file)?;
            continue;
        }
        drop(file);
        let file = archive.by_index(index)?;
        let mut bytes = Vec::new();
        file.take(MAX_OPAQUE_BYTES - total + 1)
            .read_to_end(&mut bytes)
            .map_err(|error| Error::Zip(ZipError::Io(error)))?;
        total += u64::try_from(bytes.len()).unwrap_or(u64::MAX);
        if total > MAX_OPAQUE_BYTES {
            return Ok(None);
        }
        parts.push(OpaquePart { name, bytes });
    }
    Ok(Some((parts, content.finish()?.into_inner())))
}

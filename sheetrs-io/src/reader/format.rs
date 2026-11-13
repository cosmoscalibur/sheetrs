//! The formats of workbook that are read, and how one is told from its content.

use crate::error::Result;
use std::fmt;
use std::io::Cursor;

/// The formats of workbook that are read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Format {
    Xlsx,
    Xlsm,
    Ods,
}

impl Format {
    /// The extension a file of this format has, in lowercase and without the dot.
    #[must_use]
    pub const fn extension(self) -> &'static str {
        match self {
            Self::Xlsx => "xlsx",
            Self::Xlsm => "xlsm",
            Self::Ods => "ods",
        }
    }

    /// The format of the workbook in `bytes`, read from its content (see [`Format::detect_reader`]).
    ///
    /// # Errors
    ///
    /// As [`detect_reader`](Self::detect_reader).
    pub fn detect(bytes: &[u8]) -> Result<Self> {
        Self::detect_reader(Cursor::new(bytes))
    }

    /// The format of the workbook in `reader`, read from its content and not from the name of a file:
    /// an `OpenDocument` package says `application/vnd.oasis.opendocument.spreadsheet` in its
    /// `mimetype` entry, and an Office Open XML package gives the content type of its main part in
    /// `[Content_Types].xml` (macro-enabled for an XLSM). A file called `.xlsx` that holds an ODS is
    /// an ODS.
    ///
    /// # Errors
    ///
    /// [`Error::Unsupported`](crate::Error::Unsupported) naming what was found when the input is empty, is not a zip archive (or
    /// is a truncated one), is another `OpenDocument` or Office Open XML document, an Excel template or
    /// binary workbook, or is a zip that is neither.
    pub fn detect_reader<R: std::io::Read + std::io::Seek>(reader: R) -> Result<Self> {
        super::detect::detect_reader(reader)
    }
}

impl fmt::Display for Format {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.extension())
    }
}

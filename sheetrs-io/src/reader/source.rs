//! Where a workbook comes from: a reference to it, not the workbook.
//!
//! A [`Source`] is what every tool takes as input (the linter reads a model from it, the
//! statistics stream it, the editor streams it into another file). It names a file, or holds the
//! bytes of one, and nothing is read until a tool reads it, so it is cheap to make, to clone and to
//! hand to several tools. It is not a navigable model of the workbook: that is
//! [`Workbook`], which only the tools that need every cell ask for.

use super::{
    CellSink, Selection, SheetInfo, Workbook, read_selection_from_reader_into, read_workbook,
    read_workbook_from_reader, read_workbook_from_reader_into, read_workbook_into,
};
use crate::error::{Error, Result};
use std::fmt;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::str::FromStr;
use std::sync::Arc;

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

    /// The format of a file by its extension (any case), `None` for an extension that is not one
    /// of a workbook the suite reads.
    #[must_use]
    pub fn from_path(path: &Path) -> Option<Self> {
        path.extension()?.to_str()?.parse().ok()
    }
}

impl FromStr for Format {
    type Err = Error;

    fn from_str(extension: &str) -> Result<Self> {
        match extension
            .trim_start_matches('.')
            .to_ascii_lowercase()
            .as_str()
        {
            "xlsx" => Ok(Self::Xlsx),
            "xlsm" => Ok(Self::Xlsm),
            "ods" => Ok(Self::Ods),
            _ => Err(Error::unsupported(format!(
                "'{extension}' is not an XLSX, XLSM or ODS workbook"
            ))),
        }
    }
}

impl fmt::Display for Format {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.extension())
    }
}

/// A workbook to be read: a file, or the bytes of one. It is a reference, not a model of the
/// workbook, and reads nothing until a tool does.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Source {
    /// A file, whose format is the one of its extension. Whether it exists is found out when it is
    /// read.
    File { path: PathBuf, format: Format },
    /// The bytes of a workbook of the given format.
    Bytes { data: Arc<[u8]>, format: Format },
}

impl Source {
    /// A file; it does not look at the file.
    ///
    /// # Errors
    ///
    /// When its extension is not one of a workbook the suite reads (`xlsx`, `xlsm`, `ods`).
    pub fn from_file(path: impl Into<PathBuf>) -> Result<Self> {
        let path = path.into();
        let Some(format) = Format::from_path(&path) else {
            return Err(Error::unsupported(format!(
                "'{}' is not an XLSX, XLSM or ODS workbook (by its extension)",
                path.display()
            )));
        };
        Ok(Self::File { path, format })
    }

    /// The bytes of a workbook of `format`.
    pub fn from_bytes(data: impl Into<Arc<[u8]>>, format: Format) -> Self {
        Self::Bytes {
            data: data.into(),
            format,
        }
    }

    #[must_use]
    pub fn format(&self) -> Format {
        match self {
            Self::File { format, .. } | Self::Bytes { format, .. } => *format,
        }
    }

    /// The path of the file, `None` for bytes.
    #[must_use]
    pub fn path(&self) -> Option<&Path> {
        match self {
            Self::File { path, .. } => Some(path),
            Self::Bytes { .. } => None,
        }
    }

    /// Reads the whole workbook into the model the linter analyzes.
    ///
    /// # Errors
    ///
    /// When the file cannot be opened or the archive is not a valid workbook of its format.
    pub fn read_workbook(&self) -> Result<Workbook> {
        match self {
            Self::File { path, .. } => read_workbook(path),
            Self::Bytes { data, format } => {
                read_workbook_from_reader(Cursor::new(data), Some(format.extension()))
            }
        }
    }

    /// Streams the cells of the workbook into `sink` as they are read, without keeping them: the
    /// model it returns holds what is not a cell (sheet names, defined names, sizes).
    ///
    /// # Errors
    ///
    /// As [`read_workbook`](Self::read_workbook).
    pub fn read_workbook_into(&self, sink: &mut dyn CellSink) -> Result<Workbook> {
        match self {
            Self::File { path, .. } => read_workbook_into(path, sink),
            Self::Bytes { data, format } => {
                read_workbook_from_reader_into(Cursor::new(data), Some(format.extension()), sink)
            }
        }
    }
}

impl Source {
    /// Reads what `selection` asks for and nothing else, handing the cells to `sink`: a sheet
    /// that is not selected is not parsed, the reading of a sheet stops after the last row a
    /// window reaches, and the properties not asked for are not built. Returns the sheets the
    /// reader went through (see [`WorkbookReader::read_selection_into`](super::WorkbookReader::read_selection_into)).
    ///
    /// # Errors
    ///
    /// When the file cannot be opened or the archive is not a valid workbook of its format.
    pub fn read_selection_into(
        &self,
        selection: &Selection,
        sink: &mut dyn CellSink,
    ) -> Result<Vec<SheetInfo>> {
        match self {
            Self::File { path, format } => {
                let file = std::fs::File::open(path).map_err(|source| Error::Open {
                    path: path.clone(),
                    source,
                })?;
                read_selection_from_reader_into(file, Some(format.extension()), selection, sink)
            }
            Self::Bytes { data, format } => read_selection_from_reader_into(
                Cursor::new(data),
                Some(format.extension()),
                selection,
                sink,
            ),
        }
    }
}

impl Source {
    /// Whether the workbook numbers its dates from 1904 (an ODS whose null date is 1904-01-01)
    /// instead of 1900.
    ///
    /// # Errors
    ///
    /// As [`read_workbook`](Self::read_workbook).
    pub fn uses_1904_dates(&self) -> Result<bool> {
        match self {
            Self::File { path, format } => {
                let file = std::fs::File::open(path).map_err(|source| Error::Open {
                    path: path.clone(),
                    source,
                })?;
                super::read_date1904_from_reader(file, Some(format.extension()))
            }
            Self::Bytes { data, format } => {
                super::read_date1904_from_reader(Cursor::new(data), Some(format.extension()))
            }
        }
    }
}

impl TryFrom<&Path> for Source {
    type Error = Error;

    fn try_from(path: &Path) -> Result<Self> {
        Self::from_file(path)
    }
}

impl TryFrom<PathBuf> for Source {
    type Error = Error;

    fn try_from(path: PathBuf) -> Result<Self> {
        Self::from_file(path)
    }
}

impl TryFrom<&str> for Source {
    type Error = Error;

    fn try_from(path: &str) -> Result<Self> {
        Self::from_file(path)
    }
}

impl fmt::Display for Source {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::File { path, .. } => write!(f, "{}", path.display()),
            Self::Bytes { data, format } => write!(f, "{} bytes of {format}", data.len()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const XLSX: &[u8] = include_bytes!("../../../tests/minimal_test.xlsx");

    #[test]
    fn a_format_is_read_from_an_extension_in_any_case() {
        assert_eq!("XLSX".parse::<Format>().unwrap(), Format::Xlsx);
        assert_eq!(".ods".parse::<Format>().unwrap(), Format::Ods);
        assert!("csv".parse::<Format>().is_err());
        assert_eq!(Format::from_path(Path::new("a/b.XlsM")), Some(Format::Xlsm));
        assert_eq!(Format::from_path(Path::new("a/b")), None);
        assert_eq!(Format::Ods.to_string(), "ods");
    }

    #[test]
    fn a_file_source_is_a_reference_and_reads_nothing() {
        let source = Source::from_file("does/not/exist.xlsx").unwrap();
        assert_eq!(source.format(), Format::Xlsx);
        assert_eq!(source.path(), Some(Path::new("does/not/exist.xlsx")));
        // It is found out when it is read.
        assert!(source.read_workbook().is_err());
        assert!(Source::from_file("notes.txt").is_err());
        assert_eq!(Source::try_from("a.ods").unwrap().format(), Format::Ods);
    }

    #[test]
    fn a_source_of_bytes_reads_like_a_file() {
        let source = Source::from_bytes(XLSX, Format::Xlsx);
        assert_eq!(source.path(), None);
        let book = source.read_workbook().unwrap();
        assert!(!book.sheets.is_empty());
        assert!(source.to_string().contains("bytes of xlsx"));
        assert_eq!(source.clone(), source);
    }

    #[test]
    fn the_errors_say_what_is_wrong_with_a_type() {
        use crate::Error;
        assert!(matches!(
            "csv".parse::<Format>(),
            Err(Error::Unsupported(_))
        ));
        assert!(matches!(
            Source::from_file("notes.txt"),
            Err(Error::Unsupported(_))
        ));
        // A missing file is an I/O failure that a caller can find by its source chain.
        let error = Source::from_file("does/not/exist.xlsx")
            .unwrap()
            .read_workbook()
            .unwrap_err();
        assert!(matches!(&error, Error::Open { path, .. } if path.ends_with("exist.xlsx")));
        let io = std::error::Error::source(&error)
            .and_then(|e| e.downcast_ref::<std::io::Error>())
            .expect("an I/O error is the cause");
        assert_eq!(io.kind(), std::io::ErrorKind::NotFound);
        // Bytes that are not a workbook are an archive error, not an I/O one.
        let corrupt = Source::from_bytes(&b"not a zip"[..], Format::Xlsx).read_workbook();
        assert!(matches!(corrupt, Err(Error::Archive(_))));
    }
}

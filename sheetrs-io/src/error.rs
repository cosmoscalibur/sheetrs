//! The errors of reading a workbook.

use std::path::PathBuf;

/// What can go wrong reading a workbook.
///
/// `Display` gives the whole chain (`Failed to find xl/workbook.xml: zip archive error: …`), and
/// [`std::error::Error::source`] gives the cause of each link, so a caller can tell an I/O failure
/// (by walking the sources to a [`std::io::Error`]) from a workbook that is not valid.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A file could not be opened.
    #[error("Failed to open file: {}", path.display())]
    Open {
        path: PathBuf,
        #[source]
        source: std::io::Error,
    },
    /// Reading the file failed.
    #[error(transparent)]
    Io(#[from] std::io::Error),
    /// The container of the workbook is not a valid zip archive, or a part is missing from it.
    #[error("zip archive error: {0}")]
    Archive(#[from] zip::result::ZipError),
    /// A part of the workbook is not well-formed XML.
    #[error("XML parsing error: {0}")]
    Xml(#[from] quick_xml::Error),
    /// The file is not of a format the suite reads.
    #[error("{0}")]
    Unsupported(String),
    /// The content of the workbook is not what its format says.
    #[error("{0}")]
    Malformed(String),
    /// Another error, with what was being done when it happened.
    #[error("{message}: {source}")]
    Context {
        message: String,
        #[source]
        source: Box<Error>,
    },
}

impl Error {
    pub(crate) fn malformed(message: impl Into<String>) -> Self {
        Self::Malformed(message.into())
    }

    pub(crate) fn unsupported(message: impl Into<String>) -> Self {
        Self::Unsupported(message.into())
    }
}

/// The result of reading a workbook.
pub type Result<T, E = Error> = std::result::Result<T, E>;

macro_rules! malformed_from {
    ($($error:ty),* $(,)?) => {$(
        impl From<$error> for Error {
            fn from(error: $error) -> Self {
                Self::Malformed(error.to_string())
            }
        }
    )*};
}

malformed_from!(
    quick_xml::events::attributes::AttrError,
    std::str::Utf8Error,
    std::string::FromUtf8Error,
    std::num::ParseIntError,
    std::num::ParseFloatError,
);

/// `.context("what was being done")` on a result, as the readers write it.
pub(crate) trait Context<T> {
    fn context(self, message: impl Into<String>) -> Result<T>;
}

impl<T, E: Into<Error>> Context<T> for std::result::Result<T, E> {
    fn context(self, message: impl Into<String>) -> Result<T> {
        self.map_err(|error| Error::Context {
            message: message.into(),
            source: Box::new(error.into()),
        })
    }
}

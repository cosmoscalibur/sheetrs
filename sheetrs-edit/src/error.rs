//! The errors of editing a workbook.

pub use sheetrs_io::write::ScopeError;

/// What can go wrong editing a workbook.
///
/// `Display` gives the whole chain, and [`std::error::Error::source`] the cause of each link, so a
/// caller can tell a failure of the file system (by walking the sources to a
/// [`std::io::Error`]) from a plan or a workbook that cannot be edited.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    // What was asked.
    #[error(transparent)]
    Scope(#[from] ScopeError),
    /// A plan with no operation: a copy of the file is not what was asked for.
    #[error("No operations to apply")]
    EmptyPlan,
    /// The formulas could not be calculated again.
    #[error("the formulas could not be calculated again, which '{operation}' needs: {reason}")]
    Recalculation { operation: String, reason: String },
    /// A `calculate` step that could not be run, when its value ranges were asked for (so that the mark
    /// to recalculate when the workbook is opened cannot stand for them), or that was asked for
    /// something that is not valid.
    #[error("the workbook could not be calculated: {reason}")]
    Calculation { reason: String },
    /// The value ranges of a calculation were asked for and the run kept none.
    #[error("the value ranges of a calculation were asked for, and the run kept none")]
    NoCalculation,
    /// There is no file to replace: the workbook was given as bytes.
    #[error("The workbook was given as bytes: there is no file to replace, write to a path")]
    NoFile,
    /// A search or a locale the replacement does not reproduce.
    #[error("unsupported search: {0}")]
    Search(String),

    // What is wrong with the file.
    /// What the writers of the workbook failed with (a sheet that is not there, a deletion that
    /// cannot be made right, the output being the input, a 3D reference…).
    #[error(transparent)]
    Write(#[from] sheetrs_io::write::Error),
    #[error("{0}")]
    Malformed(String),
    #[error(transparent)]
    Io(#[from] std::io::Error),
    #[error("zip archive error: {0}")]
    Archive(#[from] zip::result::ZipError),
    #[error("Error parsing XML: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error(transparent)]
    Read(#[from] sheetrs_io::Error),

    /// Another error, with what was being done when it happened.
    #[error("{message}: {source}")]
    Context {
        message: String,
        #[source]
        source: Box<Error>,
    },
}

/// The result of editing a workbook.
pub type Result<T, E = Error> = std::result::Result<T, E>;

macro_rules! from_malformed {
    ($($error:ty),* $(,)?) => {$(
        impl From<$error> for Error {
            fn from(error: $error) -> Self {
                Self::Malformed(error.to_string())
            }
        }
    )*};
}

from_malformed!(
    quick_xml::events::attributes::AttrError,
    std::str::Utf8Error,
    std::string::FromUtf8Error,
    std::num::ParseIntError,
);

/// `.context("what was being done")` on a result, as the writers write it.
pub(crate) trait Context<T> {
    fn context(self, message: impl Into<String>) -> Result<T>;
    fn with_context(self, message: impl FnOnce() -> String) -> Result<T>;
}

impl<T, E: Into<Error>> Context<T> for std::result::Result<T, E> {
    fn context(self, message: impl Into<String>) -> Result<T> {
        self.map_err(|error| Error::Context {
            message: message.into(),
            source: Box::new(error.into()),
        })
    }

    fn with_context(self, message: impl FnOnce() -> String) -> Result<T> {
        self.map_err(|error| Error::Context {
            message: message(),
            source: Box::new(error.into()),
        })
    }
}

/// `Option::context`: a value that is not there is an error with a message.
impl<T> Context<T> for Option<T> {
    fn context(self, message: impl Into<String>) -> Result<T> {
        self.ok_or_else(|| Error::Malformed(message.into()))
    }

    fn with_context(self, message: impl FnOnce() -> String) -> Result<T> {
        self.ok_or_else(|| Error::Malformed(message()))
    }
}

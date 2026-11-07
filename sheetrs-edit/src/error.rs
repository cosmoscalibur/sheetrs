//! The errors of editing a workbook.

use std::path::PathBuf;

/// A scope (`Sheet`, `Sheet!B:D`, `Sheet!3:10`) that cannot be read.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum ScopeError {
    #[error("missing sheet name in '{spec}'")]
    MissingSheet { spec: String },
    #[error("invalid columns or rows '{target}' in '{spec}'")]
    InvalidTarget { target: String, spec: String },
    #[error("unexpected '{rest}' after quoted sheet name in '{spec}'")]
    TrailingText { rest: String, spec: String },
    #[error("unterminated quoted sheet name in '{spec}'")]
    Unterminated { spec: String },
    /// Rows are deleted with a row range.
    #[error("'{spec}' is not a row range such as 'Sheet!3:10'")]
    NotRows { spec: String },
    /// Columns are deleted with a column range.
    #[error("'{spec}' is not a column range such as 'Sheet!B:D'")]
    NotColumns { spec: String },
}

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
    /// An operation addresses a sheet that an earlier one removed.
    #[error("'{operation}' addresses the sheet '{sheet}', which an earlier operation removed")]
    RemovedSheet { operation: String, sheet: String },
    /// The sheet is not in the workbook.
    #[error("Sheet '{sheet}' not found in {part}")]
    SheetNotFound { sheet: String, part: String },
    #[error("Cannot remove every sheet of the workbook")]
    EverySheet,
    /// The output would overwrite the file it is made from.
    #[error("Output file is the input file: {}", .0.display())]
    OutputIsInput(PathBuf),
    #[error("Unsupported file format")]
    UnsupportedFormat,

    // What cannot be done faithfully.
    /// A reference across sheets (`Sheet1:Sheet3!A1`): not supported yet.
    #[error(
        "the formula '{formula}' holds a 3D reference (a span of sheets such as Sheet1:Sheet3!A1), which is not supported yet when rows or columns are deleted or sheets are removed (https://github.com/cosmoscalibur/sheetrs/issues/75)"
    )]
    ThreeDReference { formula: String },
    /// `LibreOffice` places such a shape by sizes that the file does not hold.
    #[error(
        "a deletion in sheet '{sheet}' removes a cell that a shape is anchored to together with the rows or columns before it: LibreOffice places the shape by its size and the height of the rows (or width of the columns), which the file does not hold"
    )]
    ShapeDependsOnSizes { sheet: String },
    /// A change the writer cannot make so that the result is right (a deletion that covers part
    /// of an array formula, a pivot table whose source it would not adjust, …): the command
    /// stops instead of writing a workbook that is wrong.
    #[error("{0}")]
    NotAdjusted(String),

    // What is wrong with the file.
    #[error("cannot read '{formula}': {message}")]
    Formula { formula: String, message: String },
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

impl Error {
    pub(crate) fn malformed(message: impl Into<String>) -> Self {
        Self::Malformed(message.into())
    }

    pub(crate) fn not_adjusted(message: impl Into<String>) -> Self {
        Self::NotAdjusted(message.into())
    }
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

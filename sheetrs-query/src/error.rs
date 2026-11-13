//! What can go wrong reading or searching a workbook.

/// A query that cannot be answered.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The request is not valid, or names a sheet the workbook does not have.
    #[error("{0}")]
    Request(String),
    /// A search the matcher does not reproduce: a regular expression outside the subset of ICU
    /// that was compared with `LibreOffice` (see `sheetrs_engine::regex_subset`).
    #[error("unsupported search: {0}")]
    Unsupported(String),
    /// The workbook cannot be read.
    #[error(transparent)]
    Read(#[from] sheetrs_io::Error),
}

pub type Result<T> = std::result::Result<T, Error>;

impl From<sheetrs_protocol::Error> for Error {
    fn from(e: sheetrs_protocol::Error) -> Self {
        Self::Request(e.to_string())
    }
}

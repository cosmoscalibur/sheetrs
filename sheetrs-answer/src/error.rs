//! The errors of filling a template.

/// What can go wrong writing the answer of a program into its workbook.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The answer of a program cannot be written into a template (the program was not built in a
    /// workbook mode, the request is not a valid one, the run failed, or the program gave no value
    /// for a formula of the workbook).
    #[error("{0}")]
    Template(String),
    /// The writer of the workbook failed (the file cannot be read or written, a format that is not
    /// supported).
    #[error(transparent)]
    Write(#[from] sheetrs_io::write::Error),
}

/// The result of filling a template.
pub type Result<T, E = Error> = std::result::Result<T, E>;

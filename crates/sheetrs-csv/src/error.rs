use thiserror::Error;

/// A failure while reading a delimited text file.
#[derive(Debug, Error)]
pub enum Error {
    /// The input cannot be read.
    #[error("cannot read the input: {0}")]
    Io(#[from] std::io::Error),
    /// The text holds more than the cell store can address.
    #[error("workbook too large: {0}")]
    TooLarge(#[from] sheetrs_core::LimitError),
}

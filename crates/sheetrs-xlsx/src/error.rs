use thiserror::Error;

/// A failure while reading an XLSX package.
#[derive(Debug, Error)]
pub enum Error {
    /// The input is not a readable ZIP package.
    #[error("not a readable ZIP package: {0}")]
    Zip(#[from] zip::result::ZipError),
    /// A part the workbook refers to is not in the package.
    #[error("missing package part {0:?}")]
    MissingPart(String),
    /// A part is not well-formed XML.
    #[error("malformed XML in {part}: {source}")]
    Xml {
        /// The package part being read.
        part: String,
        /// The parser error.
        source: quick_xml::Error,
    },
    /// A part is well-formed XML that does not describe a valid workbook.
    #[error("invalid {part}: {reason}")]
    Malformed {
        /// The package part being read.
        part: String,
        /// What is wrong with it.
        reason: String,
    },
}

impl Error {
    pub(crate) fn xml(part: &str, source: quick_xml::Error) -> Self {
        Self::Xml {
            part: part.to_owned(),
            source,
        }
    }

    pub(crate) fn malformed(part: &str, reason: impl Into<String>) -> Self {
        Self::Malformed {
            part: part.to_owned(),
            reason: reason.into(),
        }
    }
}

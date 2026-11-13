//! The errors of the engine: what can go wrong running a program, reading its request, or
//! packing it into an executable.

use std::io;
use std::path::PathBuf;

/// What went wrong.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A request that cannot be answered: an input or a range that does not exist, a value of the
    /// wrong kind, a program that fails. The text is meant to be shown to whoever sent it.
    #[error("{0}")]
    Request(String),
    /// A command line that is wrong: a flag without a value, a value that does not parse, an
    /// option the document defines.
    #[error("{0}")]
    Usage(String),
    /// The JSON of a request is not JSON.
    #[error("invalid JSON input: {0}")]
    Json(#[from] serde_json::Error),
    /// A file that cannot be read or written.
    #[error("{}: {source}", path.display())]
    Io {
        path: PathBuf,
        #[source]
        source: io::Error,
    },
    /// The request cannot be read from the standard input.
    #[error("cannot read the request: {0}")]
    Stdin(#[source] io::Error),
    /// A program that cannot be encoded into the payload of an executable.
    #[error("cannot encode the program: {0}")]
    Encode(#[from] postcard::Error),
    /// A block of constants that does not fit the payload format.
    #[error("a block of constants is too big")]
    BlockTooBig,
    /// A constant of a program that lies on a sheet the program does not list.
    #[error("a constant lies on a sheet the program does not list")]
    ConstantOnUnknownSheet,
}

impl Error {
    pub(crate) fn io(path: impl Into<PathBuf>, source: io::Error) -> Self {
        Self::Io {
            path: path.into(),
            source,
        }
    }
}

/// The result of an operation of the engine.
pub type Result<T, E = Error> = std::result::Result<T, E>;

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
    /// A range that was asked for and that the program cannot read: it lies outside its outputs.
    /// When the outputs were declared (a settings sheet or a settings file) `declared` lists them, as
    /// `name (range)`, and a range can only be one of them or lie inside one; without a declaration
    /// the outputs are the formula cells of the workbook, and `declared` is `None`.
    #[error("{}", outside_message(range, declared.as_deref()))]
    OutsideOutputs {
        /// The range as it was asked for.
        range: String,
        /// The declared outputs, `name (range)` each.
        declared: Option<Vec<String>>,
    },
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

/// What a range outside the outputs of a program is told: the declared outputs, when there are some.
fn outside_message(range: &str, declared: Option<&[String]>) -> String {
    match declared {
        Some(outputs) => format!(
            "`{range}` is outside the outputs this program declares: a range can only be one of them or lie inside one ({}; see --meta)",
            outputs.join(", ")
        ),
        None => format!(
            "`{range}` cannot be read: it is not inside the outputs of this program, which are the formula cells of the workbook as none were declared (see --meta)"
        ),
    }
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

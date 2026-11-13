//! The errors of the compiler: what can go wrong reading a workbook and its configuration, and
//! compiling it.

use std::io;
use std::path::PathBuf;

/// What went wrong.
///
/// The variants that carry a message (`Config`, `Settings`, `Locale`, `Compile`) are diagnostics
/// about the input of a build, written for the person who wrote it (they name the sheet, the row
/// or the key at fault); the other ones keep the error they come from.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A file that cannot be read or written.
    #[error("{}: {source}", path.display())]
    Io {
        path: PathBuf,
        #[source]
        source: io::Error,
    },
    /// A workbook that cannot be read.
    #[cfg(feature = "load")]
    #[error("{what}: {source:#}")]
    Workbook {
        /// The workbook, as its [`Book`](sheetrs_io::Book) names it.
        what: String,
        #[source]
        source: sheetrs_io::Error,
    },
    /// A build configuration that is not valid TOML for the configuration.
    #[error("{origin}: {source}")]
    Toml {
        origin: String,
        #[source]
        source: toml::de::Error,
    },
    /// A build configuration with a value that is not valid.
    #[error("{0}")]
    Config(String),
    /// A name in `enable_features` that is not a feature.
    #[error("unknown feature `{name}` in enable_features (known: {})", known.join(", "))]
    UnknownFeature {
        name: String,
        known: &'static [&'static str],
    },
    /// A settings sheet that does not declare what a build needs.
    #[error("{0}")]
    Settings(String),
    /// A locale that is not valid or not complete.
    #[error("{0}")]
    Locale(String),
    /// A workbook that cannot be compiled, or a range that cannot be resolved against it.
    #[error("{0}")]
    Compile(String),
    /// `SOURCE_DATE_EPOCH` set to something that is not a number of seconds.
    #[error("SOURCE_DATE_EPOCH must be a number of seconds since 1970, not `{0}`")]
    BuildTime(String),
    /// The recalculation by `LibreOffice` that did not happen.
    #[error("{0}")]
    Recalculation(String),
    /// A WebAssembly module that cannot be completed with a program (see [`crate::wasm_target`]).
    #[error("{0}")]
    Wasm(String),
    /// A failure of the engine (packing the executable, a run).
    #[error(transparent)]
    Engine(#[from] sheetrs_engine::Error),
    /// The answer that could not be written into the workbook.
    #[cfg(feature = "load")]
    #[error("{0:#}")]
    Answer(#[from] sheetrs_answer::Error),
}

impl Error {
    pub(crate) fn io(path: impl Into<PathBuf>, source: io::Error) -> Self {
        Self::Io {
            path: path.into(),
            source,
        }
    }
}

/// The result of an operation of the compiler.
pub type Result<T, E = Error> = std::result::Result<T, E>;

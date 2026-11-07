//! What can be wrong with a request.

/// A request that cannot be understood. The text is meant to be shown to whoever sent it.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The request is not valid: a field that does not exist, a value of the wrong kind.
    #[error("{0}")]
    Request(String),
    /// The text is not JSON, or does not have the shape of the request.
    #[error("invalid JSON: {0}")]
    Json(String),
}

pub type Result<T> = std::result::Result<T, Error>;

impl From<serde_json::Error> for Error {
    fn from(e: serde_json::Error) -> Self {
        Self::Json(e.to_string())
    }
}

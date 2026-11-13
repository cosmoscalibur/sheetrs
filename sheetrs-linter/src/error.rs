//! The errors of the linter's configuration.

use std::path::PathBuf;

/// What can be wrong with a configuration. The text of the validation errors names what to fix.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The configuration file could not be read.
    #[error("Failed to read the configuration {}: {source}", path.display())]
    File {
        path: PathBuf,
        #[source]
        source: std::io::Error,
    },
    /// The text is not TOML of the shape of a configuration.
    #[error("invalid configuration: {0}")]
    Syntax(#[from] toml::de::Error),
    /// `ALL` disables every rule, which is not what a configuration is for.
    #[error("Configuration error: 'ALL' is not allowed in global disabled_rules")]
    AllDisabled,
    /// A rule or category that the linter does not have (a misspelled rule would otherwise be
    /// ignored without a word).
    #[error("Configuration error: Unknown rule or category '{token}' in {place}")]
    UnknownRule { token: String, place: String },
    /// A parameter that must be a whole number of at least 1.
    #[error(
        "Configuration error: '{key}' in {scope} must be a whole number of at least 1, got {value}"
    )]
    NotPositive {
        key: &'static str,
        scope: String,
        value: String,
    },
    /// `enable_features` is not an array of names.
    #[error("Configuration error: '{key}' must be an array of feature names ({known})")]
    FeatureList { key: &'static str, known: String },
    /// A compatibility feature that does not exist.
    #[error("Configuration error: unknown feature '{name}' in '{key}' (known: {known})")]
    UnknownFeature {
        name: String,
        key: &'static str,
        known: String,
    },
}

/// The result of building a configuration.
pub type Result<T, E = Error> = std::result::Result<T, E>;

//! What a run reports besides its result, as structured records.
//!
//! Every tool that has something to say that is not an error of the run (a calculation the engine
//! could not do, a note about what was built) says it as a [`Diagnostic`]: a [`Level`], a stable
//! `code` that a program can match on, the `message` for a person and the `context` fields the
//! message was made from (`operation`, `sheet`, `cell`, `reason`...). The text is for people; the
//! code and the fields are for programs. Each layer shows them in its own way: the command line as a
//! line of text or of JSON on the standard error ([`Logger`]), Python through `logging` and
//! `warnings`, WebAssembly as objects handed to a callback, Rust as the values themselves.

use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;
use std::fmt;
use std::str::FromStr;

/// How serious a diagnostic is, from the one that always matters to the one for whoever is
/// following the run.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
#[non_exhaustive]
pub enum Level {
    /// The run failed.
    Error,
    /// The run went through and something that was asked was not done as asked.
    Warning,
    /// What the run did (what was written, what was built).
    Info,
    /// Detail for whoever looks for a cause.
    Debug,
}

impl Level {
    /// The name of the level, as it is written in the records.
    #[must_use]
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Error => "error",
            Self::Warning => "warning",
            Self::Info => "info",
            Self::Debug => "debug",
        }
    }
}

impl fmt::Display for Level {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

impl FromStr for Level {
    type Err = String;

    fn from_str(text: &str) -> Result<Self, String> {
        match text {
            "error" => Ok(Self::Error),
            "warning" => Ok(Self::Warning),
            "info" => Ok(Self::Info),
            "debug" => Ok(Self::Debug),
            other => Err(format!(
                "a level is `error`, `warning`, `info` or `debug`, not `{other}`"
            )),
        }
    }
}

/// One thing a run reports besides its result: a level, a stable code, a message and the context
/// fields the message was made from.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Diagnostic {
    /// How serious it is.
    pub level: Level,
    /// The stable name of what happened (`RECALC_NOT_SUPPORTED`): it does not change with the
    /// wording of the message.
    pub code: String,
    /// What happened, for a person.
    pub message: String,
    /// The fields the message was made from, by name (`operation`, `sheet`, `cell`, `reason`...).
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub context: BTreeMap<String, String>,
}

impl Diagnostic {
    /// A diagnostic without context fields.
    #[must_use]
    pub fn new(level: Level, code: impl Into<String>, message: impl Into<String>) -> Self {
        Self {
            level,
            code: code.into(),
            message: message.into(),
            context: BTreeMap::new(),
        }
    }

    /// A warning.
    #[must_use]
    pub fn warning(code: impl Into<String>, message: impl Into<String>) -> Self {
        Self::new(Level::Warning, code, message)
    }

    /// An informative record.
    #[must_use]
    pub fn info(code: impl Into<String>, message: impl Into<String>) -> Self {
        Self::new(Level::Info, code, message)
    }

    /// An error.
    #[must_use]
    pub fn error(code: impl Into<String>, message: impl Into<String>) -> Self {
        Self::new(Level::Error, code, message)
    }

    /// The same diagnostic with one more context field.
    #[must_use]
    pub fn with(mut self, field: impl Into<String>, value: impl Into<String>) -> Self {
        self.context.insert(field.into(), value.into());
        self
    }

    /// The record as JSON: `{"level", "code", "message", "context"}` (no `context` without fields).
    #[must_use]
    pub fn to_json(&self) -> serde_json::Value {
        serde_json::to_value(self).unwrap_or(serde_json::Value::Null)
    }
}

impl fmt::Display for Diagnostic {
    /// The line the command line writes: `warning: message`; an informative record is the message
    /// alone.
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.level {
            Level::Info => f.write_str(&self.message),
            other => write!(f, "{other}: {}", self.message),
        }
    }
}

/// How the command line writes a diagnostic on the standard error.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum LogFormat {
    /// A line of text (`warning: message`).
    #[default]
    Human,
    /// A JSON record per line.
    Json,
}

impl FromStr for LogFormat {
    type Err = String;

    fn from_str(text: &str) -> Result<Self, String> {
        match text {
            "human" => Ok(Self::Human),
            "json" => Ok(Self::Json),
            other => Err(format!("a log format is `human` or `json`, not `{other}`")),
        }
    }
}

/// Where the command line sends the diagnostics: the standard error, with the diagnostics at or above
/// the threshold, as text or as JSON.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Logger {
    threshold: Level,
    format: LogFormat,
}

impl Default for Logger {
    /// Warnings and errors, as text.
    fn default() -> Self {
        Self {
            threshold: Level::Warning,
            format: LogFormat::Human,
        }
    }
}

impl Logger {
    /// A logger for the flags of the command line: `quiet` keeps only the errors and each `verbose`
    /// adds a level (`info`, then `debug`).
    #[must_use]
    pub fn new(quiet: bool, verbose: u8, format: LogFormat) -> Self {
        let threshold = if quiet {
            Level::Error
        } else {
            match verbose {
                0 => Level::Warning,
                1 => Level::Info,
                _ => Level::Debug,
            }
        };
        Self { threshold, format }
    }

    /// Whether a diagnostic of this level is written.
    #[must_use]
    pub fn enabled(&self, level: Level) -> bool {
        level <= self.threshold
    }

    /// The line a diagnostic is written as, or `None` when its level is below the threshold.
    #[must_use]
    pub fn line(&self, diagnostic: &Diagnostic) -> Option<String> {
        self.enabled(diagnostic.level).then(|| match self.format {
            LogFormat::Human => diagnostic.to_string(),
            LogFormat::Json => diagnostic.to_json().to_string(),
        })
    }

    /// Writes a diagnostic on the standard error, if its level is enabled.
    pub fn emit(&self, diagnostic: &Diagnostic) {
        if let Some(line) = self.line(diagnostic) {
            eprintln!("{line}");
        }
    }

    /// Writes each diagnostic (see [`Logger::emit`]).
    pub fn emit_all<'a>(&self, diagnostics: impl IntoIterator<Item = &'a Diagnostic>) {
        for diagnostic in diagnostics {
            self.emit(diagnostic);
        }
    }
}

/// The flags every command line has for its diagnostics (feature `cli`).
#[cfg(feature = "cli")]
#[derive(Debug, Clone, Copy, Default, clap::Args)]
pub struct LogArgs {
    /// How diagnostics are written on the standard error: `human` (text, the default) or `json` (a
    /// record per line with `level`, `code`, `message` and `context`)
    #[arg(long, value_name = "FORMAT", default_value = "human", global = true, value_parser = parse_format)]
    pub log_format: LogFormat,

    /// Write only the errors
    #[arg(short, long, global = true, conflicts_with = "verbose")]
    pub quiet: bool,

    /// Also write what the run did (`-v`) and the detail (`-vv`)
    #[arg(short, long, action = clap::ArgAction::Count, global = true)]
    pub verbose: u8,
}

#[cfg(feature = "cli")]
fn parse_format(text: &str) -> Result<LogFormat, String> {
    text.parse()
}

#[cfg(feature = "cli")]
impl LogArgs {
    /// The logger these flags ask for.
    #[must_use]
    pub fn logger(&self) -> Logger {
        Logger::new(self.quiet, self.verbose, self.log_format)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn warning() -> Diagnostic {
        Diagnostic::warning("RECALC_NOT_SUPPORTED", "the workbook was not calculated")
            .with("operation", "calculate")
            .with("reason", "XLOOKUP")
    }

    #[test]
    fn a_diagnostic_is_a_record_with_a_stable_code_and_its_fields() {
        assert_eq!(
            warning().to_json(),
            serde_json::json!({
                "level": "warning",
                "code": "RECALC_NOT_SUPPORTED",
                "message": "the workbook was not calculated",
                "context": {"operation": "calculate", "reason": "XLOOKUP"},
            })
        );
        let plain = Diagnostic::error("X", "m");
        assert!(plain.to_json().get("context").is_none());
        let back: Diagnostic = serde_json::from_value(warning().to_json()).unwrap();
        assert_eq!(back, warning());
    }

    #[test]
    fn the_command_line_writes_the_levels_the_flags_ask_for() {
        let info = Diagnostic::info("WRITTEN", "wrote out.xlsx");
        let error = Diagnostic::error("FAILED", "no");
        let line = |logger: Logger, d: &Diagnostic| logger.line(d);

        let default = Logger::default();
        assert_eq!(line(default, &info), None);
        assert_eq!(
            line(default, &warning()).as_deref(),
            Some("warning: the workbook was not calculated")
        );

        let quiet = Logger::new(true, 0, LogFormat::Human);
        assert_eq!(line(quiet, &warning()), None);
        assert_eq!(line(quiet, &error).as_deref(), Some("error: no"));

        let verbose = Logger::new(false, 1, LogFormat::Human);
        assert_eq!(line(verbose, &info).as_deref(), Some("wrote out.xlsx"));
        let debug = Diagnostic::new(Level::Debug, "D", "detail");
        assert_eq!(line(verbose, &debug), None);
        assert_eq!(
            line(Logger::new(false, 2, LogFormat::Human), &debug).as_deref(),
            Some("debug: detail")
        );
    }

    #[test]
    fn the_json_format_is_one_record_per_line() {
        let json = Logger::new(false, 0, LogFormat::Json);
        let line = json.line(&warning()).unwrap();
        assert!(!line.contains('\n'));
        let record: serde_json::Value = serde_json::from_str(&line).unwrap();
        assert_eq!(record["code"], "RECALC_NOT_SUPPORTED");
        assert_eq!(record["context"]["reason"], "XLOOKUP");
    }
}

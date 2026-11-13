//! The three ways a run is written, behind one entry point: [`write_report`] for a person
//! (`human`), for an interface of your own (`json`) and for the tools
//! that read the standard ([`sarif`](crate::sarif)). The executable, the WebAssembly and Python
//! bindings all go through it.

use crate::config::LinterConfig;
use crate::violation::Violation;
use sheetrs_io::reader::Workbook;
use std::fmt;
use std::io::{self, Write};
use std::str::FromStr;

/// How a run is written.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum OutputFormat {
    /// The report for a person, as plain text or with terminal styles.
    Human,
    /// The JSON document of `sheetrs lint --format json`.
    Json,
    /// A SARIF 2.1.0 log.
    Sarif,
}

impl OutputFormat {
    /// The name the executable and the bindings use: `human`, `json` or `sarif`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Human => "human",
            Self::Json => "json",
            Self::Sarif => "sarif",
        }
    }
}

impl fmt::Display for OutputFormat {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

impl FromStr for OutputFormat {
    type Err = String;

    fn from_str(name: &str) -> Result<Self, Self::Err> {
        match name {
            "human" => Ok(Self::Human),
            "json" => Ok(Self::Json),
            "sarif" => Ok(Self::Sarif),
            other => Err(format!(
                "unknown format `{other}`: use `human`, `json` or `sarif`"
            )),
        }
    }
}

/// Writes `violations` of the workbook `file` (the path as the caller gave it) to `out` in
/// `format`. `config` is the configuration the linter ran with (SARIF lists its rules); `color`
/// asks the human report for terminal styles and changes nothing else.
///
/// # Errors
///
/// A write error of `out`.
pub fn write_report(
    out: &mut impl Write,
    format: OutputFormat,
    file: &str,
    violations: &[Violation],
    workbook: &Workbook,
    config: &LinterConfig,
    color: bool,
) -> io::Result<()> {
    match format {
        OutputFormat::Human => crate::human::write_human(out, file, violations, workbook, color),
        OutputFormat::Json => {
            crate::json::write_json(out, file, violations, workbook).map_err(io::Error::from)
        }
        OutputFormat::Sarif => crate::sarif::write_sarif(out, file, violations, workbook, config)
            .map_err(io::Error::from),
    }
}

/// [`write_report`] into a string.
///
/// # Errors
///
/// Never in practice: the writers produce UTF-8 and a `Vec` does not fail.
pub fn render_report(
    format: OutputFormat,
    file: &str,
    violations: &[Violation],
    workbook: &Workbook,
    config: &LinterConfig,
    color: bool,
) -> io::Result<String> {
    let mut out = Vec::new();
    write_report(&mut out, format, file, violations, workbook, config, color)?;
    String::from_utf8(out).map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_names_round_trip() {
        for format in [OutputFormat::Human, OutputFormat::Json, OutputFormat::Sarif] {
            assert_eq!(format.to_string().parse::<OutputFormat>(), Ok(format));
        }
        assert!("xml".parse::<OutputFormat>().is_err());
    }

    #[test]
    fn every_format_writes_a_run_without_violations() {
        let workbook = Workbook::default();
        let config = LinterConfig::default();
        let human = render_report(
            OutputFormat::Human,
            "b.xlsx",
            &[],
            &workbook,
            &config,
            false,
        )
        .unwrap();
        assert!(human.contains("No violations found"));
        let json =
            render_report(OutputFormat::Json, "b.xlsx", &[], &workbook, &config, false).unwrap();
        let json: serde_json::Value = serde_json::from_str(&json).unwrap();
        assert_eq!(json["summary"]["total"], 0);
        let sarif = render_report(
            OutputFormat::Sarif,
            "b.xlsx",
            &[],
            &workbook,
            &config,
            false,
        )
        .unwrap();
        let sarif: serde_json::Value = serde_json::from_str(&sarif).unwrap();
        assert_eq!(sarif["version"], "2.1.0");
    }
}

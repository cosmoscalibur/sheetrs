//! The `sheetrs lint` subcommand: the arguments, and the choice of the standard output for the
//! report that the library renders.

use crate::output::write_report;
use crate::{Linter, LinterConfig, Severity, Violation};
use anyhow::{Context, Result};
use clap::ValueEnum;
use sheetrs_io::Book;
use sheetrs_io::reader::Workbook;
use std::io::{BufWriter, Write};
use std::path::{Path, PathBuf};
use std::process::ExitCode;

/// The arguments of `sheetrs lint`.
#[derive(clap::Args)]
#[command(about = "Fast Excel/ODS linter with hierarchical violation reporting")]
pub struct Args {
    /// The Excel/ODS workbook to lint
    #[arg(
        short = 'i',
        long = "input",
        visible_alias = "file",
        value_name = "FILE"
    )]
    pub file: PathBuf,

    /// Path to configuration file (TOML)
    #[arg(short, long, value_name = "CONFIG")]
    pub config: Option<PathBuf>,

    /// Output format
    #[arg(short, long, value_enum, default_value = "human")]
    pub format: OutputFormat,

    /// Show only errors (hide warnings and info)
    #[arg(short, long)]
    pub errors_only: bool,
}

/// How the report is written.
#[derive(Clone, ValueEnum)]
pub enum OutputFormat {
    /// Human-readable colored output
    Human,
    /// JSON document, for an interface of your own
    Json,
    /// SARIF 2.1.0 log, for code scanning, editors and CI
    Sarif,
}

impl From<&OutputFormat> for crate::OutputFormat {
    fn from(format: &OutputFormat) -> Self {
        match format {
            OutputFormat::Human => Self::Human,
            OutputFormat::Json => Self::Json,
            OutputFormat::Sarif => Self::Sarif,
        }
    }
}

/// Runs `sheetrs lint`: the exit code is 1 when a violation is an error, and 0 otherwise.
///
/// # Errors
///
/// When the configuration or the workbook cannot be read, or the report cannot be written.
#[allow(clippy::needless_pass_by_value)] // the subcommand hands over its arguments, as `main` gets them
pub fn run(cli: Args) -> Result<ExitCode> {
    // Load configuration
    let config = if let Some(config_path) = &cli.config {
        LinterConfig::from_file(config_path)
            .with_context(|| format!("Failed to load config from {}", config_path.display()))?
    } else {
        // Try to load default config from current directory if it exists
        let default_config_path = PathBuf::from("sheetrs.toml");
        if default_config_path.exists() {
            LinterConfig::from_file(&default_config_path).with_context(|| {
                format!(
                    "Failed to load config from {}",
                    default_config_path.display()
                )
            })?
        } else {
            LinterConfig::default()
        }
    };

    let linter = Linter::with_config(config);

    let workbook = Book::from_file(&cli.file)
        .and_then(|source| source.workbook())
        .with_context(|| format!("Failed to lint file: {}", cli.file.display()))?;
    let violations = linter.lint_workbook(&workbook);

    // Filter violations if needed
    let violations: Vec<_> = if cli.errors_only {
        violations
            .into_iter()
            .filter(|v| v.severity == Severity::Error)
            .collect()
    } else {
        violations
    };

    // Output results
    print(
        (&cli.format).into(),
        &cli.file,
        &violations,
        &workbook,
        linter.config(),
    )?;

    // Exit with appropriate code
    let exit_code = if violations.is_empty() {
        0
    } else {
        i32::from(violations.iter().any(|v| v.severity == Severity::Error))
    };

    Ok(ExitCode::from(u8::try_from(exit_code).unwrap_or(1)))
}

/// Writes the violations to the standard output in `format`, with terminal styles for the human
/// report (the `colored` crate drops them when the output is not a terminal).
///
/// # Errors
///
/// A write error of the standard output.
fn print(
    format: crate::OutputFormat,
    file_path: &Path,
    violations: &[Violation],
    workbook: &Workbook,
    config: &LinterConfig,
) -> anyhow::Result<()> {
    let mut out = BufWriter::new(std::io::stdout().lock());
    write_report(
        &mut out,
        format,
        &file_path.display().to_string(),
        violations,
        workbook,
        config,
        true,
    )?;
    out.flush()?;
    Ok(())
}

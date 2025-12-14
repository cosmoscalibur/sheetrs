use anyhow::{Context, Result};
use clap::{Parser, ValueEnum};
use sheetcraft_core::{Linter, LinterConfig, Severity};
use std::path::PathBuf;

mod formatter;

#[derive(Parser)]
#[command(name = "sheetlint")]
#[command(about = "Fast Excel/ODS linter with hierarchical violation reporting", long_about = None)]
#[command(version)]
struct Cli {
    /// Path to the Excel/ODS file to lint
    #[arg(value_name = "FILE")]
    file: PathBuf,

    /// Path to configuration file (TOML)
    #[arg(short, long, value_name = "CONFIG")]
    config: Option<PathBuf>,

    /// Output format
    #[arg(short, long, value_enum, default_value = "human")]
    format: OutputFormat,

    /// Show only errors (hide warnings and info)
    #[arg(short, long)]
    errors_only: bool,
}

#[derive(Clone, ValueEnum)]
enum OutputFormat {
    /// Human-readable colored output
    Human,
    /// JSON output for CI/CD integration
    Json,
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    // Load configuration
    let config = if let Some(config_path) = &cli.config {
        LinterConfig::from_file(config_path)
            .with_context(|| format!("Failed to load config from {}", config_path.display()))?
    } else {
        // Try to load default config from current directory if it exists
        let default_config_path = PathBuf::from("sheetlint.toml");
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

    // Validate configuration
    let valid_tokens = sheetcraft_core::rules::registry::get_all_valid_tokens();
    config
        .validate_rules(&valid_tokens)
        .context("Invalid configuration")?;

    // Create linter and run
    let linter = Linter::with_config(config);

    let violations = linter
        .lint_file(&cli.file)
        .with_context(|| format!("Failed to lint file: {}", cli.file.display()))?;

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
    match cli.format {
        OutputFormat::Human => {
            formatter::print_human(&cli.file, &violations);
        }
        OutputFormat::Json => {
            formatter::print_json(&cli.file, &violations)?;
        }
    }

    // Exit with appropriate code
    let exit_code = if violations.is_empty() {
        0
    } else if violations.iter().any(|v| v.severity == Severity::Error) {
        1
    } else {
        0 // Only warnings, still exit 0
    };

    std::process::exit(exit_code);
}

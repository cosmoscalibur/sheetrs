//! The `sheetrs stats` subcommand: the statistics of a workbook, as text or as JSON.

use crate::BookStats as _;
use crate::{WorkbookStats, percentage};
use anyhow::{Context, Result};
use clap::ValueEnum;
use sheetrs_io::Book;
use std::path::PathBuf;
use std::process::ExitCode;

/// The arguments of `sheetrs stats`.
#[derive(clap::Args)]
#[command(
    about = "Statistics of a workbook: cells, formulas, errors, comments, charts and tables, by sheet"
)]
pub struct Args {
    /// The Excel/ODS workbook to read
    #[arg(
        short = 'i',
        long = "input",
        visible_alias = "file",
        value_name = "FILE"
    )]
    pub file: PathBuf,

    /// Output format
    #[arg(short, long, value_enum, default_value = "human")]
    pub format: OutputFormat,
}

/// How the statistics are written.
#[derive(Clone, ValueEnum)]
pub enum OutputFormat {
    /// Human-readable output
    Human,
    /// JSON output
    Json,
}

/// Runs `sheetrs stats`.
///
/// # Errors
///
/// When the workbook cannot be read or the output cannot be written.
#[allow(clippy::needless_pass_by_value)] // the subcommand hands over its arguments, as `main` gets them
pub fn run(cli: Args) -> Result<ExitCode> {
    let stats = Book::from_file(&cli.file)
        .and_then(|book| book.stats())
        .with_context(|| format!("Failed to read file: {}", cli.file.display()))?;

    match cli.format {
        OutputFormat::Human => print_human(&stats),
        OutputFormat::Json => print_json(&stats)?,
    }

    Ok(ExitCode::SUCCESS)
}

fn humanize_size(bytes: u64) -> String {
    const KB: u64 = 1024;
    const MB: u64 = KB * 1024;
    const GB: u64 = MB * 1024;

    if bytes >= GB {
        format!("{:.2} GB", bytes as f64 / GB as f64)
    } else if bytes >= MB {
        format!("{:.2} MB", bytes as f64 / MB as f64)
    } else if bytes >= KB {
        format!("{:.2} KB", bytes as f64 / KB as f64)
    } else {
        format!("{bytes} bytes")
    }
}

/// Format a count with its share of the workbook total.
fn with_share(part: u64, total: u64) -> String {
    format!("{part} ({:.2}%)", percentage(part, total))
}

fn print_human(stats: &WorkbookStats) {
    println!("Workbook Statistics:");
    println!("  Total Sheets: {}", stats.total_sheets);
    println!("  Total Named Ranges: {}", stats.total_named_ranges);
    println!(
        "  Total File Size: {}",
        humanize_size(stats.total_file_size)
    );
    println!("  Total Used Cells: {}", stats.total_used_cells);
    println!("  Total Formulas: {}", stats.total_formulas);
    println!("  Total Errors: {}", stats.total_errors);
    println!("  Total Empty Cells: {}", stats.total_empty_cells);
    println!("  Total Comments: {}", stats.total_comments);
    println!("  Total Charts: {}", stats.total_charts);
    println!("  Total Tables: {}", stats.total_tables);
    println!(
        "  Total Sheet XML Size: {}",
        humanize_size(stats.total_xml_size)
    );

    for sheet in &stats.sheets {
        println!("\nSheet: {}", sheet.name);
        println!(
            "  Used Cells: {}",
            with_share(sheet.used_cells, stats.total_used_cells)
        );
        println!(
            "  Formulas: {}",
            with_share(sheet.formulas, stats.total_formulas)
        );
        println!("  Errors: {}", with_share(sheet.errors, stats.total_errors));
        println!(
            "  Empty Cells: {}",
            with_share(sheet.empty_cells, stats.total_empty_cells)
        );
        println!(
            "  Used Range: {}",
            sheet.used_range.as_deref().unwrap_or("-")
        );
        println!(
            "  End Of Sheet: {}",
            sheet.end_of_sheet.as_deref().unwrap_or("-")
        );
        println!("  Comments: {}", sheet.comments);
        println!("  Charts: {}", sheet.charts);
        println!("  Tables: {}", sheet.tables);
        println!(
            "  XML Size: {} ({:.2}%)",
            humanize_size(sheet.xml_size),
            percentage(sheet.xml_size, stats.total_xml_size)
        );
    }
}

fn print_json(stats: &WorkbookStats) -> Result<()> {
    let json = serde_json::to_string_pretty(stats)?;
    println!("{json}");
    Ok(())
}

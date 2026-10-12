#![warn(dead_code_pub_in_binary)]

mod stats;

use anyhow::{Context, Result};
use clap::{Parser, ValueEnum};
use sheetrs::reader;
use std::path::PathBuf;

use stats::{WorkbookStats, percentage};

#[derive(Parser)]
#[command(name = "sheetstats")]
#[command(about = "Statistics generator for SheetRS")]
#[command(version)]
struct Cli {
    /// Path to the Excel/ODS file
    #[arg(value_name = "FILE")]
    file: PathBuf,

    /// Output format
    #[arg(short, long, value_enum, default_value = "human")]
    format: OutputFormat,
}

#[derive(Clone, ValueEnum)]
enum OutputFormat {
    /// Human-readable output
    Human,
    /// JSON output
    Json,
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    let workbook = reader::read_workbook(&cli.file)
        .with_context(|| format!("Failed to read file: {}", cli.file.display()))?;
    let stats = WorkbookStats::from_workbook(&workbook);

    match cli.format {
        OutputFormat::Human => print_human(&stats),
        OutputFormat::Json => print_json(&stats)?,
    }

    Ok(())
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

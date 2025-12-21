use anyhow::{Context, Result};
use clap::Parser;
use sheetrs::writer;
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "sheetcli")]
#[command(about = "CLI tools for SheetChecks", long_about = None)]
#[command(version)]
struct Cli {
    /// Path to the Excel/ODS file
    #[arg(value_name = "FILE")]
    file: PathBuf,

    /// Remove specified sheets
    #[arg(long, num_args = 1.., value_name = "SHEET")]
    remove_sheets: Vec<String>,

    /// Remove specified named ranges
    #[arg(long, num_args = 1.., value_name = "RANGE")]
    remove_ranges: Vec<String>,

    /// Output file (Required for destructive operations)
    #[arg(short, long)]
    output: Option<PathBuf>,

    /// Show what would be done without making changes
    #[arg(long)]
    dry_run: bool,
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    let has_ops = !cli.remove_sheets.is_empty() || !cli.remove_ranges.is_empty();

    if !has_ops {
        println!("No operations specified. Use --help for usage.");
        return Ok(());
    }

    // Enforce output file for destructive operations
    if cli.output.is_none() {
        anyhow::bail!("Output file is required for destructive operations. Use --output <FILE>.");
    }
    let output_path = cli.output.unwrap();

    // Prepare modifications
    use std::collections::HashSet;
    let mut mods = writer::WorkbookModifications::default();

    if !cli.remove_sheets.is_empty() {
        mods.remove_sheets = Some(cli.remove_sheets.iter().cloned().collect::<HashSet<_>>());
    }
    if !cli.remove_ranges.is_empty() {
        mods.remove_named_ranges = Some(cli.remove_ranges.iter().cloned().collect::<HashSet<_>>());
    }

    if cli.dry_run {
        println!("[DRY RUN] Operations on '{}':", cli.file.display());
        if let Some(sheets) = &mods.remove_sheets {
            println!("  Removing sheets:");
            for s in sheets {
                println!("    - {}", s);
            }
        }
        if let Some(ranges) = &mods.remove_named_ranges {
            println!("  Removing named ranges:");
            for r in ranges {
                println!("    - {}", r);
            }
        }
        println!("\nOutput would be: {}", output_path.display());
    } else {
        println!("Modifying '{}'...", cli.file.display());
        writer::modify_workbook(&cli.file, &output_path, &mods)
            .with_context(|| "Failed to modify workbook")?;

        println!("✓ Successfully modified workbook");
        println!("Output: {}", output_path.display());
    }

    Ok(())
}

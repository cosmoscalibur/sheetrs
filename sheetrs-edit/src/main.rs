#![warn(dead_code_pub_in_binary)]

use anyhow::{Context, Result};
use clap::Parser;
use sheetrs_core::reader::Source;
use sheetrs_edit::{ColumnRange, EditPlan, Operation, RowRange, writer};
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "sheetedit")]
#[command(about = "CLI tools for SheetRS", long_about = None)]
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

    /// Replace formulas by their cached values in a sheet (`Sheet`), columns
    /// (`Sheet!B:D`) or rows (`Sheet!3:10`)
    #[arg(long, num_args = 1.., value_name = "SCOPE")]
    convert_to_values: Vec<writer::ValueScope>,

    /// Delete rows (`Sheet!3:10`, or `Sheet!3` for one), in the rows of the input file. The
    /// references to them become `#REF!` and the ranges around them shrink, as in `LibreOffice`
    #[arg(long, num_args = 1.., value_name = "ROWS")]
    delete_rows: Vec<RowRange>,

    /// Delete columns (`Sheet!B:D`, or `Sheet!B` for one), in the columns of the input file
    #[arg(long, num_args = 1.., value_name = "COLUMNS")]
    delete_cols: Vec<ColumnRange>,

    /// List all named ranges with their values and status
    #[arg(long)]
    list_ranges: bool,

    /// Output file (Required for destructive operations)
    #[arg(short, long)]
    output: Option<PathBuf>,

    /// Show what would be done without making changes
    #[arg(long)]
    dry_run: bool,
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    let has_ops = !cli.remove_sheets.is_empty()
        || !cli.remove_ranges.is_empty()
        || !cli.convert_to_values.is_empty()
        || !cli.delete_rows.is_empty()
        || !cli.delete_cols.is_empty()
        || cli.list_ranges;

    if !has_ops {
        println!("No operations specified. Use --help for usage.");
        return Ok(());
    }

    if cli.list_ranges {
        println!("Listing named ranges for '{}':", cli.file.display());
        let workbook = sheetrs_core::reader::read_workbook(&cli.file)
            .with_context(|| format!("Failed to read file: {}", cli.file.display()))?;

        println!("Found {} defined names", workbook.defined_names.len());

        if workbook.defined_names.is_empty() {
            println!("  No named ranges found.");
        } else {
            // Sort for consistent output
            let mut names: Vec<_> = workbook.defined_names.iter().collect();
            names.sort_by_key(|(k, _)| k.as_str());

            for (name, value) in names {
                // ERR002 check: Broken reference contains #REF!
                let is_broken = value.contains("#REF!");
                let status = if is_broken { "❌" } else { "✅" };
                println!("  {status} {name} -> {value}");
            }
        }
        return Ok(());
    }

    // Enforce output file for destructive operations
    if cli.output.is_none() {
        anyhow::bail!("Output file is required for destructive operations. Use --output <FILE>.");
    }
    let output_path = cli.output.unwrap();

    // The flags of a command are all written against the input file, in no order: `against_input`
    // makes of them a plan of one pass, however many flags and in whatever order they came.
    let plan = EditPlan::against_input([
        Operation::RemoveSheets(cli.remove_sheets),
        Operation::RemoveNamedRanges(cli.remove_ranges),
        Operation::ConvertToValues(cli.convert_to_values),
        Operation::DeleteRows(cli.delete_rows),
        Operation::DeleteColumns(cli.delete_cols),
    ]);
    if cli.dry_run {
        println!("[DRY RUN] Operations on '{}':", cli.file.display());
        print!("{}", plan.explain()?);
        println!("\nOutput would be: {}", output_path.display());
    } else {
        println!("Modifying '{}'...", cli.file.display());
        plan.execute(&Source::from_file(&cli.file)?, &output_path)
            .with_context(|| "Failed to modify workbook")?;

        println!("✓ Successfully modified workbook");
        println!("Output: {}", output_path.display());
    }

    Ok(())
}

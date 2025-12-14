use anyhow::{Context, Result};
use clap::{Parser, ValueEnum};
use serde::Serialize;
use sheetcraft_core::reader;
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "sheetstats")]
#[command(about = "Statistics generator for SheetChecks")]
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

#[derive(Serialize)]
struct FileStats {
    total_sheets: usize,
    total_named_ranges: usize,
    sheet_sizes: Vec<SheetSize>,
    total_file_size: u64,
    formula_stats: Vec<FormulaStats>,
    cell_stats: Vec<CellStats>,
}

#[derive(Serialize)]
struct SheetSize {
    name: String,
    compressed_size: u64,
    percentage: f64,
}

#[derive(Serialize)]
struct FormulaStats {
    sheet_name: String,
    formula_count: usize,
    percentage: f64,
}

#[derive(Serialize)]
struct CellStats {
    sheet_name: String,
    non_empty_count: usize,
    percentage: f64,
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    // Read workbook
    let workbook = reader::read_workbook(&cli.file)
        .with_context(|| format!("Failed to read file: {}", cli.file.display()))?;

    // Get basic stats
    let total_sheets = workbook.sheets.len();
    let total_named_ranges = workbook.defined_names.len();

    // Get file size
    let total_file_size = std::fs::metadata(&cli.file)
        .with_context(|| "Failed to get file size")?
        .len();

    // Calculate sheet sizes (only for XLSX files)
    let sheet_sizes = if cli.file.extension().and_then(|s| s.to_str()) == Some("xlsx") {
        calculate_sheet_sizes(&cli.file, total_file_size, &workbook)?
    } else {
        // For ODS or other formats, we can't easily calculate per-sheet sizes
        vec![]
    };

    // Calculate formula statistics
    let formula_stats = calculate_formula_stats(&workbook);

    // Calculate cell statistics
    let cell_stats = calculate_cell_stats(&workbook);

    let stats = FileStats {
        total_sheets,
        total_named_ranges,
        sheet_sizes,
        total_file_size,
        formula_stats,
        cell_stats,
    };

    // Output results
    match cli.format {
        OutputFormat::Human => print_human(&stats),
        OutputFormat::Json => print_json(&stats)?,
    }

    Ok(())
}

fn calculate_formula_stats(workbook: &sheetcraft_core::reader::Workbook) -> Vec<FormulaStats> {
    let mut total_formulas = 0usize;
    let mut sheet_formulas: Vec<(String, usize)> = Vec::new();

    for sheet in &workbook.sheets {
        let formula_count = sheet
            .all_cells()
            .filter(|cell| cell.value.is_formula())
            .count();

        total_formulas += formula_count;
        sheet_formulas.push((sheet.name.clone(), formula_count));
    }

    // Calculate percentages
    sheet_formulas
        .into_iter()
        .map(|(name, count)| FormulaStats {
            sheet_name: name,
            formula_count: count,
            percentage: if total_formulas > 0 {
                (count as f64 / total_formulas as f64) * 100.0
            } else {
                0.0
            },
        })
        .collect()
}

fn calculate_cell_stats(workbook: &sheetcraft_core::reader::Workbook) -> Vec<CellStats> {
    let mut total_non_empty = 0usize;
    let mut sheet_cells: Vec<(String, usize)> = Vec::new();

    for sheet in &workbook.sheets {
        let non_empty_count = sheet.all_cells().count();

        total_non_empty += non_empty_count;
        sheet_cells.push((sheet.name.clone(), non_empty_count));
    }

    // Calculate percentages
    sheet_cells
        .into_iter()
        .map(|(name, count)| CellStats {
            sheet_name: name,
            non_empty_count: count,
            percentage: if total_non_empty > 0 {
                (count as f64 / total_non_empty as f64) * 100.0
            } else {
                0.0
            },
        })
        .collect()
}

fn calculate_sheet_sizes(
    file_path: &PathBuf,
    total_size: u64,
    workbook: &sheetcraft_core::reader::Workbook,
) -> Result<Vec<SheetSize>> {
    use std::fs::File;
    use std::io::BufReader;
    use zip::ZipArchive;

    let file = File::open(file_path)?;
    let reader = BufReader::new(file);
    let mut archive = ZipArchive::new(reader)?;

    let mut sheet_sizes = Vec::new();

    // Iterate through ZIP entries to find sheet XML files
    for i in 0..archive.len() {
        let file = archive.by_index(i)?;
        let name = file.name().to_string();

        // Sheet files are in xl/worksheets/sheetN.xml
        if name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml") {
            let compressed_size = file.compressed_size();
            let percentage = (compressed_size as f64 / total_size as f64) * 100.0;

            // Extract sheet number from filename (1-indexed)
            let sheet_num: usize = name
                .trim_start_matches("xl/worksheets/sheet")
                .trim_end_matches(".xml")
                .parse()
                .unwrap_or(0);

            // Get actual sheet name from workbook (0-indexed)
            let sheet_name = if sheet_num > 0 && sheet_num <= workbook.sheets.len() {
                workbook.sheets[sheet_num - 1].name.clone()
            } else {
                format!("Sheet{}", sheet_num)
            };

            sheet_sizes.push(SheetSize {
                name: sheet_name,
                compressed_size,
                percentage,
            });
        }
    }

    Ok(sheet_sizes)
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
        format!("{} bytes", bytes)
    }
}

fn print_human(stats: &FileStats) {
    println!("File Statistics:");
    println!("  Total Sheets: {}", stats.total_sheets);
    println!("  Total Named Ranges: {}", stats.total_named_ranges);
    println!(
        "  Total File Size: {}",
        humanize_size(stats.total_file_size)
    );

    if !stats.sheet_sizes.is_empty() {
        println!("\nSheet Sizes (compressed):");
        for sheet in &stats.sheet_sizes {
            println!(
                "  {}: {} ({:.2}%)",
                sheet.name,
                humanize_size(sheet.compressed_size),
                sheet.percentage
            );
        }
    }

    if !stats.formula_stats.is_empty() {
        println!("\nFormulas by Sheet:");
        for stat in &stats.formula_stats {
            println!(
                "  {}: {} formulas ({:.2}%)",
                stat.sheet_name, stat.formula_count, stat.percentage
            );
        }
    }

    if !stats.cell_stats.is_empty() {
        println!("\nNon-Empty Cells by Sheet:");
        for stat in &stats.cell_stats {
            println!(
                "  {}: {} cells ({:.2}%)",
                stat.sheet_name, stat.non_empty_count, stat.percentage
            );
        }
    }
}

fn print_json(stats: &FileStats) -> Result<()> {
    let json = serde_json::to_string_pretty(stats)?;
    println!("{}", json);
    Ok(())
}

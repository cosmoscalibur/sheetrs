#![warn(dead_code_pub_in_binary)]

use anyhow::{Context, Result};
use clap::Parser;
use sheetrs_edit::{ColumnRange, Editor, RowRange, ValueScope};
use sheetrs_io::reader::Source;
use sheetrs_protocol::{Find, Mode, Reference, Replace, Scope, SearchIn};
use std::path::PathBuf;

#[derive(Clone, Copy, clap::ValueEnum)]
enum Where {
    Value,
    Formula,
}

/// A scope of `--scope`: a range when it has a sheet and cells (`Data!A1:C9`, `A1:B2`), a sheet
/// otherwise.
fn parse_scope(text: &str) -> Scope {
    match Reference::parse(text) {
        Ok(reference) if text.contains('!') || reference.sheet.is_none() => Scope::Range(reference),
        _ => Scope::Sheet(text.trim_matches('\'').to_string()),
    }
}

#[derive(Parser)]
#[allow(clippy::struct_excessive_bools)] // the flags of a command line
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
    convert_to_values: Vec<ValueScope>,

    /// Delete rows (`Sheet!3:10`, or `Sheet!3` for one), in the rows of the input file. The
    /// references to them become `#REF!` and the ranges around them shrink, as in `LibreOffice`
    #[arg(long, num_args = 1.., value_name = "ROWS")]
    delete_rows: Vec<RowRange>,

    /// Delete columns (`Sheet!B:D`, or `Sheet!B` for one), in the columns of the input file
    #[arg(long, num_args = 1.., value_name = "COLUMNS")]
    delete_cols: Vec<ColumnRange>,

    /// Find and replace text in the cells, as `LibreOffice`'s Find & Replace does (needs `--with`).
    /// The text of a number is the one of its input line (`1234.5`, `25.6%`, `03/15/2024`); the
    /// replacement is entered as if it were typed (`12` makes a number, `=1+1` a formula, nothing an
    /// empty cell)
    #[arg(long, value_name = "TEXT", requires = "with")]
    replace: Option<String>,

    /// What takes the place of what `--replace` finds (`$1` and `&` in a regular expression)
    #[arg(
        long,
        value_name = "TEXT",
        requires = "replace",
        allow_hyphen_values = true
    )]
    with: Option<String>,

    /// `--replace` is a regular expression (the subset of ICU that is reproduced)
    #[arg(long, requires = "replace", conflicts_with = "wildcards")]
    regex: bool,

    /// `--replace` has wildcards (`*`, `?`, `\` to escape)
    #[arg(long, requires = "replace")]
    wildcards: bool,

    /// Distinguish upper and lower case
    #[arg(long, requires = "replace")]
    match_case: bool,

    /// The text must be the whole content of the cell
    #[arg(long, requires = "replace")]
    entire_cells: bool,

    /// Where `--replace` looks: the value (default) or the formula, in the notation of `LibreOffice`
    #[arg(long = "in", value_enum, requires = "replace", default_value = "value")]
    search_in: Where,

    /// Where `--replace` applies: a sheet (`Data`) or a range (`Data!A1:C9`); repeatable, the whole
    /// workbook when omitted
    #[arg(long, value_name = "SCOPE", requires = "replace")]
    scope: Vec<String>,

    /// List all named ranges with their values and status
    #[arg(long)]
    list_ranges: bool,

    /// Output file (required unless `--in-place`)
    #[arg(short, long, conflicts_with = "in_place")]
    output: Option<PathBuf>,

    /// Replace the input file with the result (written to a temporary file and renamed, so a
    /// failure leaves the input as it was)
    #[arg(long)]
    in_place: bool,

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
        || cli.replace.is_some()
        || cli.list_ranges;

    if !has_ops {
        println!("No operations specified. Use --help for usage.");
        return Ok(());
    }

    if cli.list_ranges {
        println!("Listing named ranges for '{}':", cli.file.display());
        let workbook = sheetrs_io::reader::read_workbook(&cli.file)
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

    // The output is explicit: a path, or the input file itself.
    if cli.output.is_none() && !cli.in_place {
        anyhow::bail!(
            "Output file is required for destructive operations. Use --output <FILE>, or --in-place to replace the input."
        );
    }
    let output_path = cli.output.clone().unwrap_or_else(|| cli.file.clone());

    // The flags of a command are all written against the input file, in no order. Applied in this
    // order (conversions, rows, columns, names, sheets) they are one pass, however many flags
    // there are and in whatever order they came.
    let source = Source::from_file(&cli.file)?;
    let mut editor = Editor::new(&source);
    // A replacement reads the cells: it goes first, against the input file like the rest.
    if let (Some(text), Some(with)) = (&cli.replace, &cli.with) {
        editor.replace(
            cli.scope.iter().map(|s| parse_scope(s)),
            Replace {
                find: Find {
                    text: text.clone(),
                    mode: if cli.regex {
                        Mode::Regex
                    } else if cli.wildcards {
                        Mode::Wildcards
                    } else {
                        Mode::Literal
                    },
                    match_case: cli.match_case,
                    entire_cells: cli.entire_cells,
                    search_in: match cli.search_in {
                        Where::Value => SearchIn::Value,
                        Where::Formula => SearchIn::Formula,
                    },
                    include_formulas: true,
                },
                with: with.clone(),
            },
        );
    }
    editor
        .convert_to_values(cli.convert_to_values)
        .delete_rows(cli.delete_rows)
        .delete_columns(cli.delete_cols)
        .remove_named_ranges(cli.remove_ranges)
        .remove_sheets(cli.remove_sheets);
    if cli.dry_run {
        println!("[DRY RUN] Operations on '{}':", cli.file.display());
        print!("{}", editor.explain()?);
        println!("\nOutput would be: {}", output_path.display());
    } else {
        println!("Modifying '{}'...", cli.file.display());
        let report = if cli.in_place {
            editor
                .overwrite()
                .with_context(|| "Failed to modify workbook")?
        } else {
            editor
                .save(&output_path)
                .with_context(|| "Failed to modify workbook")?
        };

        println!("✓ Successfully modified workbook");
        println!("Output: {}", output_path.display());
        if cli.replace.is_some() {
            println!("Replaced in {} cells", report.replaced);
            for skipped in &report.skipped {
                println!(
                    "  not replaced: {}!{}: {}",
                    skipped.sheet, skipped.cell, skipped.reason
                );
            }
        }
    }

    Ok(())
}

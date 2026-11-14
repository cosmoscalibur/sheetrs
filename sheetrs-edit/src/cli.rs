//! The `sheetrs edit` subcommand: the edits of a workbook, written as one pass.

use crate::{BookEdit, ColumnRange, RowRange, ValueScope};
use anyhow::{Context, Result};
use sheetrs_io::Book;
use sheetrs_protocol::{Find, Logger, Mode, Replace, Scope, SearchIn};
use std::path::PathBuf;
use std::process::ExitCode;

/// What a replacement looks at.
#[derive(Clone, Copy, clap::ValueEnum)]
pub enum Where {
    Value,
    Formula,
    Comment,
}

/// The arguments of `sheetrs edit`.
#[derive(clap::Args)]
#[allow(clippy::struct_excessive_bools)] // the flags of a command line
#[command(
    about = "Edit a workbook: delete or hide sheets, rows and columns, delete names, turn formulas into values, find and replace"
)]
pub struct Args {
    /// The Excel/ODS workbook to edit
    #[arg(
        short = 'i',
        long = "input",
        visible_alias = "file",
        value_name = "FILE"
    )]
    pub file: PathBuf,

    /// Delete these sheets
    #[arg(long, num_args = 1.., value_name = "SHEET")]
    pub delete_sheets: Vec<String>,

    /// Delete these defined names (named ranges)
    #[arg(long, num_args = 1.., value_name = "RANGE")]
    pub delete_names: Vec<String>,

    /// Replace formulas by their cached values in a sheet (`Sheet`), columns
    /// (`Sheet!B:D`) or rows (`Sheet!3:10`)
    #[arg(long, num_args = 1.., value_name = "SCOPE")]
    pub formula_to_value: Vec<ValueScope>,

    /// Delete rows (`Sheet!3:10`, or `Sheet!3` for one), in the rows of the input file. The
    /// references to them become `#REF!` and the ranges around them shrink, as in `LibreOffice`
    #[arg(long, num_args = 1.., value_name = "ROWS")]
    pub delete_rows: Vec<RowRange>,

    /// Delete columns (`Sheet!B:D`, or `Sheet!B` for one), in the columns of the input file
    #[arg(long, num_args = 1.., value_name = "COLUMNS")]
    pub delete_columns: Vec<ColumnRange>,

    /// Hide rows (`Sheet!3:10`), keeping them in the file. In the rows of the input file, before any
    /// deletion
    #[arg(long, num_args = 1.., value_name = "ROWS")]
    pub hide_rows: Vec<RowRange>,

    /// Show rows, whatever hid them (by hand or by a filter). Applied after `--hide-rows`
    #[arg(long, num_args = 1.., value_name = "ROWS")]
    pub show_rows: Vec<RowRange>,

    /// Hide columns (`Sheet!B:D`)
    #[arg(long, num_args = 1.., value_name = "COLUMNS")]
    pub hide_columns: Vec<ColumnRange>,

    /// Show columns. Applied after `--hide-columns`
    #[arg(long, num_args = 1.., value_name = "COLUMNS")]
    pub show_columns: Vec<ColumnRange>,

    /// Hide these sheets (one must stay visible)
    #[arg(long, num_args = 1.., value_name = "SHEET")]
    pub hide_sheets: Vec<String>,

    /// Show these sheets. Applied after `--hide-sheets`
    #[arg(long, num_args = 1.., value_name = "SHEET")]
    pub show_sheets: Vec<String>,

    /// Find and replace text in the cells, as `LibreOffice`'s Find & Replace does (needs `--with`).
    /// The text of a number is the one of its input line (`1234.5`, `25.6%`, `03/15/2024`); the
    /// replacement is entered as if it were typed (`12` makes a number, `=1+1` a formula, nothing an
    /// empty cell). Repeat it for a sweep: `--replace A --with B --replace C --with D` pairs each
    /// `--replace` with the `--with` in the same position, and every pair is an operation, in
    /// order, that sees what the previous one left; the options below apply to all of them. An empty
    /// text matches nothing: no cell is replaced
    #[arg(long, value_name = "TEXT", requires = "with")]
    pub replace: Vec<String>,

    /// What takes the place of what `--replace` finds (`$1` and `&` in a regular expression); one
    /// per `--replace`, in the same order
    #[arg(
        long,
        value_name = "TEXT",
        requires = "replace",
        allow_hyphen_values = true
    )]
    pub with: Vec<String>,

    /// `--replace` is a regular expression (ICU, the subset that is reproduced; `^` and `$` anchor the text of the cell, and the pattern may match the empty text: that match is replaced only at the end of the text)
    #[arg(long, requires = "replace", conflicts_with = "wildcards")]
    pub regex: bool,

    /// `--replace` has wildcards (`*`, `?`, `\` to escape)
    #[arg(long, requires = "replace")]
    pub wildcards: bool,

    /// Distinguish upper and lower case
    #[arg(long, requires = "replace")]
    pub match_case: bool,

    /// The text must be the whole content of the cell
    #[arg(long, requires = "replace")]
    pub entire_cells: bool,

    /// Search and replace in the text each cell shows (`1,234.50`, `15 Mar 2023`), not in its input line (`LibreOffice`'s "Search formatted")
    #[arg(long, requires = "replace")]
    pub formatted: bool,

    /// Also replace in the rows an autofilter has hidden (`LibreOffice`'s "Search filtered"); by default
    /// they are left alone, as there. Rows hidden by hand are always replaced
    #[arg(long, requires = "replace")]
    pub filtered: bool,

    /// Where `--replace` looks: the value (default), the formula in the notation of `LibreOffice`, or the text of the notes
    #[arg(long = "in", value_enum, requires = "replace", default_value = "value")]
    pub search_in: Where,

    /// Where `--replace` applies: a sheet (`Data`) or a range (`Data!A1:C9`); repeatable, the whole
    /// workbook when omitted
    #[arg(long, value_name = "SCOPE", requires = "replace")]
    pub scope: Vec<String>,

    /// Locale of the recalculation that an edit leaves to do (`TEXT`, the conversion of a number to
    /// text): a tag such as `es-CO`. Default: the `locale` of `--config`, else the one of the system
    /// (`LC_ALL`, `LC_CTYPE`, then `LANG`; `en-US` when it has none), for an XLSX and an ODS alike
    #[arg(long, value_name = "TAG")]
    pub locale: Option<String>,

    /// Serial number that `NOW()` and `TODAY()` give in the recalculation that an edit leaves to do
    /// (the `now` of `sheetc` and of a `calculate` step); default: the `now` of `--config`, else the
    /// system clock
    #[arg(long, value_name = "SERIAL")]
    pub now: Option<f64>,

    /// Seed of `RAND()` and its relatives in that recalculation (the `seed` of `sheetc` and of a
    /// `calculate` step); default: the `seed` of `--config`, else seeded from the clock
    #[arg(long, value_name = "N")]
    pub seed: Option<u64>,

    /// Configuration file (TOML) whose `locale`, `now` and `seed` in `[global]` are the defaults of
    /// `--locale`, `--now` and `--seed` (default: `sheetrs.toml` in the current directory, when there
    /// is one)
    #[arg(short = 'c', long, value_name = "FILE")]
    pub config: Option<PathBuf>,

    /// Output file (required unless `--in-place`)
    #[arg(short, long, conflicts_with = "in_place")]
    pub output: Option<PathBuf>,

    /// Replace the input file with the result (written to a temporary file and renamed, so a
    /// failure leaves the input as it was)
    #[arg(long)]
    pub in_place: bool,

    /// Show what would be done without making changes (needs no `--output` or `--in-place`; with
    /// one of them it shows where the result would go and still writes nothing)
    #[arg(long)]
    pub dry_run: bool,
}

/// Runs `sheetrs edit`.
///
/// # Errors
///
/// When the arguments are not consistent, the workbook cannot be read or the result cannot be written.
#[allow(clippy::needless_pass_by_value)] // the subcommand hands over its arguments, as `main` gets them
pub fn run(cli: Args, logger: &Logger) -> Result<ExitCode> {
    let has_ops = !cli.delete_sheets.is_empty()
        || !cli.delete_names.is_empty()
        || !cli.formula_to_value.is_empty()
        || !cli.delete_rows.is_empty()
        || !cli.delete_columns.is_empty()
        || !cli.hide_rows.is_empty()
        || !cli.show_rows.is_empty()
        || !cli.hide_columns.is_empty()
        || !cli.show_columns.is_empty()
        || !cli.hide_sheets.is_empty()
        || !cli.show_sheets.is_empty()
        || !cli.replace.is_empty();

    if !has_ops {
        println!("No operations specified. Use --help for usage.");
        return Ok(ExitCode::SUCCESS);
    }

    if cli.replace.len() != cli.with.len() {
        anyhow::bail!(
            "{} --replace but {} --with: each --replace needs its own --with, paired by position",
            cli.replace.len(),
            cli.with.len()
        );
    }

    // The output is explicit: a path, or the input file itself. A preview (`--dry-run`) writes
    // nothing, so it needs none.
    if cli.output.is_none() && !cli.in_place && !cli.dry_run {
        anyhow::bail!(
            "Output file is required for destructive operations. Use --output <FILE>, or --in-place to replace the input."
        );
    }
    let output_path = cli.output.clone().unwrap_or_else(|| cli.file.clone());

    // The flags of a command are all written against the input file, in no order: the planner
    // decides what runs first (see `sheetrs_edit::BookEdit`), as it does for every other caller.
    let mut book = Book::from_file(&cli.file)?;
    if let Some(tag) =
        sheetrs_io::locale::for_command(cli.locale.as_deref(), cli.config.as_deref())?
    {
        book = book.with_locale(tag);
    }
    book = book.with_volatile(
        sheetrs_io::Volatile {
            now: cli.now,
            seed: cli.seed,
        }
        .or(sheetrs_io::Volatile::configured(cli.config.as_deref())?),
    );
    if !cli.replace.is_empty() {
        let mode = if cli.regex {
            Mode::Regex
        } else if cli.wildcards {
            Mode::Wildcards
        } else {
            Mode::Literal
        };
        let search_in = match cli.search_in {
            Where::Value => SearchIn::Value,
            Where::Formula => SearchIn::Formula,
            Where::Comment => SearchIn::Comment,
        };
        book.replace(
            cli.scope.iter().map(|s| Scope::from(s.as_str())),
            cli.replace
                .iter()
                .zip(&cli.with)
                .map(|(text, with)| Replace {
                    find: Find {
                        text: text.clone(),
                        mode,
                        match_case: cli.match_case,
                        entire_cells: cli.entire_cells,
                        search_in,
                        include_formulas: true,
                        formatted: cli.formatted,
                        filtered: cli.filtered,
                    },
                    with: with.clone(),
                }),
        );
    }
    book.hide_rows(cli.hide_rows)
        .show_rows(cli.show_rows)
        .hide_columns(cli.hide_columns)
        .show_columns(cli.show_columns)
        .hide_sheets(cli.hide_sheets)
        .show_sheets(cli.show_sheets)
        .formula_to_value(cli.formula_to_value)
        .delete_rows(cli.delete_rows)
        .delete_columns(cli.delete_columns)
        .delete_names(cli.delete_names)
        .delete_sheets(cli.delete_sheets);
    if cli.dry_run {
        println!("[DRY RUN] Operations on '{}':", cli.file.display());
        print!("{}", book.explain());
        if cli.output.is_some() || cli.in_place {
            println!("\nOutput would be: {}", output_path.display());
        } else {
            println!("\nNothing is written: no --output or --in-place was given.");
        }
    } else {
        println!("Modifying '{}'...", cli.file.display());
        let report = if cli.in_place {
            book.write_in_place()
                .with_context(|| "Failed to modify workbook")?
        } else {
            book.write(&output_path)
                .with_context(|| "Failed to modify workbook")?
        };

        logger.emit_all(&report.diagnostics);
        println!("✓ Successfully modified workbook");
        println!("Output: {}", output_path.display());
        if !cli.replace.is_empty() {
            println!("Replaced in {} cells", report.replaced);
            for skipped in &report.skipped {
                println!(
                    "  not replaced: {}!{}: {}",
                    skipped.sheet, skipped.cell, skipped.reason
                );
            }
        }
    }

    Ok(ExitCode::SUCCESS)
}

//! `sheetrs`: one executable for the tools over a workbook. Each subcommand is the `cli` module of
//! its crate; the format of the file (XLSX or ODS) is detected from its content.

use clap::{Parser, Subcommand};
use std::process::ExitCode;

#[derive(Parser)]
#[command(
    name = "sheetrs",
    version,
    about = "Query, edit, lint and measure XLSX and ODS workbooks"
)]
struct Cli {
    #[command(flatten)]
    log: sheetrs_protocol::LogArgs,

    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    /// Read the cells of a workbook, search them, list its sheets and names
    Query(sheetrs_query::cli::Args),
    /// Edit a workbook: delete or hide, convert formulas to values, find and replace
    Edit(Box<sheetrs_edit::cli::Args>),
    /// Check a workbook against the lint rules
    Lint(sheetrs_linter::cli::Args),
    /// Statistics of a workbook
    Stats(sheetrs_stats::cli::Args),
}

fn main() -> anyhow::Result<ExitCode> {
    let cli = Cli::parse();
    let logger = cli.log.logger();
    match cli.command {
        Command::Query(args) => sheetrs_query::cli::run(args),
        Command::Edit(args) => sheetrs_edit::cli::run(*args, &logger),
        Command::Lint(args) => sheetrs_linter::cli::run(args),
        Command::Stats(args) => sheetrs_stats::cli::run(args),
    }
}

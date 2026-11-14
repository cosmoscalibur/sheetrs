//! The `sheetrs query` subcommand: the cells of a workbook by range, sheet or workbook, the
//! search of Find & Replace, and the list of its sheets and defined names.

use crate::BookQuery as _;
use anyhow::{Context, Result, bail};
use clap::ValueEnum;
use sheetrs_io::Book;
use sheetrs_io::reader::{Cell, CellSink, Properties, Selection, Window};
use sheetrs_protocol::{Find, Mode, Property, Query, Reference, Scope, SearchIn};
use std::collections::BTreeMap;
use std::io::Write;
use std::path::PathBuf;
use std::process::ExitCode;

/// A property of a cell.
#[derive(Clone, Copy, ValueEnum)]
pub enum Prop {
    Value,
    Formula,
    #[value(alias = "number_format")]
    NumberFormat,
    Display,
    Comment,
}

/// What a search looks at.
#[derive(Clone, Copy, ValueEnum)]
pub enum Where {
    Value,
    Formula,
    Comment,
}

/// The arguments of `sheetrs query`.
#[derive(clap::Args)]
#[command(
    about = "Read the cells of an XLSX or ODS workbook by range, sheet or workbook, with the properties asked for, find the cells that match a search, and list its sheets and defined names"
)]
#[allow(clippy::struct_excessive_bools)] // the flags of a command line
pub struct Args {
    /// The XLSX/ODS workbook to read
    #[arg(
        short = 'i',
        long = "input",
        visible_alias = "file",
        value_name = "FILE"
    )]
    pub file: PathBuf,

    /// A range to read (`Sheet1!A1:C9`); without a sheet, the first one. Repeatable
    #[arg(long, value_name = "RANGE")]
    pub range: Vec<String>,

    /// A sheet to read whole. Repeatable
    #[arg(long, value_name = "SHEET")]
    pub sheet: Vec<String>,

    /// Read every sheet (the default when no range or sheet is given)
    #[arg(long)]
    pub workbook: bool,

    /// The properties of each cell; the value when omitted. Comma separated or repeated
    #[arg(
        short,
        long,
        value_enum,
        value_delimiter = ',',
        value_name = "PROPERTY"
    )]
    pub property: Vec<Prop>,

    /// Keep only the cells where this text is found (as `LibreOffice`'s Find & Replace finds it); the
    /// empty text matches nothing
    #[arg(long, value_name = "TEXT")]
    pub find: Option<String>,

    /// The search is a regular expression (ICU, the subset that is reproduced; `^` and `$` anchor the text of the cell)
    #[arg(long, conflicts_with = "wildcards")]
    pub regex: bool,

    /// The search has wildcards (`*`, `?`, `\` to escape)
    #[arg(long)]
    pub wildcards: bool,

    /// Distinguish upper and lower case in the search
    #[arg(long)]
    pub match_case: bool,

    /// The search must match the whole content of the cell
    #[arg(long)]
    pub entire_cells: bool,

    /// What the search looks at
    #[arg(long = "in", value_enum, default_value = "value")]
    pub search_in: Where,

    /// A search by value also looks at the result of the formulas
    #[arg(long)]
    pub include_formulas: bool,

    /// Search the text each cell shows (`1,234.50`, `15 Mar 2023`), not its input line
    #[arg(long)]
    pub formatted: bool,

    /// Also search the rows an autofilter has hidden (`LibreOffice`'s "Search filtered"); by default
    /// they are left out, as there. Rows hidden by hand are always searched
    #[arg(long)]
    pub filtered: bool,

    /// Locale of `display` and of the dates a search sees (a tag such as `es-CO`). Default: the
    /// `locale` of `--config`, else the one of the system (`LC_ALL`, `LC_CTYPE`, then `LANG`;
    /// `en-US` when it has none), for an XLSX and an ODS alike
    #[arg(long, value_name = "TAG")]
    pub locale: Option<String>,

    /// Configuration file (TOML) whose `locale` in `[global]` is the default of `--locale`
    /// (default: `sheetrs.toml` in the current directory, when there is one)
    #[arg(short = 'c', long, value_name = "FILE")]
    pub config: Option<PathBuf>,

    /// Give the numbers whose format is a date or a time as `{"type": "date" | "time" | "datetime",
    /// "value": "<ISO 8601>"}`, worked out with the date system of the file (the serial number the file
    /// stores is the default)
    #[arg(long)]
    pub dates: bool,

    /// List the sheets of the workbook, in order, as a JSON array of their names (nothing else is read)
    #[arg(long, conflicts_with_all = ["list_names", "find", "query"])]
    pub list_sheets: bool,

    /// List the defined names of the workbook as a JSON object of name to definition
    #[arg(long, conflicts_with_all = ["find", "query"])]
    pub list_names: bool,

    /// The whole query as JSON (a text, or `@file`); replaces the options above
    #[arg(long, value_name = "JSON")]
    pub query: Option<String>,

    /// One JSON object per cell, one per line
    #[arg(long)]
    pub jsonl: bool,

    /// Indent the JSON
    #[arg(long)]
    pub pretty: bool,
}

fn build(cli: &Args) -> Result<Query> {
    if let Some(q) = &cli.query {
        let text = match q.strip_prefix('@') {
            Some(path) => std::fs::read_to_string(path).with_context(|| path.to_string())?,
            None => q.clone(),
        };
        return serde_json::from_str(&text).context("invalid query");
    }
    let mut scope = Vec::new();
    for r in &cli.range {
        scope.push(Scope::Range(
            r.parse::<Reference>().map_err(|e| anyhow::anyhow!("{e}"))?,
        ));
    }
    for s in &cli.sheet {
        scope.push(Scope::Sheet(s.clone()));
    }
    if cli.workbook || scope.is_empty() {
        scope.push(Scope::Workbook);
    }
    let properties = if cli.property.is_empty() {
        vec![Property::Value]
    } else {
        cli.property
            .iter()
            .map(|p| match p {
                Prop::Value => Property::Value,
                Prop::Formula => Property::Formula,
                Prop::NumberFormat => Property::NumberFormat,
                Prop::Display => Property::Display,
                Prop::Comment => Property::Comment,
            })
            .collect()
    };
    let find = cli.find.as_ref().map(|text| Find {
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
            Where::Comment => SearchIn::Comment,
        },
        include_formulas: cli.include_formulas,
        formatted: cli.formatted,
        filtered: cli.filtered,
    });
    if find.is_none()
        && (cli.regex
            || cli.wildcards
            || cli.match_case
            || cli.entire_cells
            || cli.include_formulas
            || cli.formatted
            || cli.filtered)
    {
        bail!("the search options need --find");
    }
    Ok(Query {
        scope,
        properties,
        find,
        locale: cli.locale.clone(),
        dates: cli.dates,
    })
}

/// Runs `sheetrs query`.
///
/// # Errors
///
/// When the query is not valid, the workbook cannot be read or the answer cannot be written.
#[allow(clippy::needless_pass_by_value)] // the subcommand hands over its arguments, as `main` gets them
pub fn run(cli: Args) -> Result<ExitCode> {
    if cli.list_sheets || cli.list_names {
        return list(&cli);
    }
    let query = build(&cli)?;
    let source = book(&cli)?;
    let response = source.query(&query)?;
    let out = std::io::stdout();
    let mut out = std::io::BufWriter::new(out.lock());
    if cli.jsonl {
        for cell in &response.cells {
            serde_json::to_writer(&mut out, cell)?;
            writeln!(out)?;
        }
    } else if cli.pretty {
        serde_json::to_writer_pretty(&mut out, &response)?;
        writeln!(out)?;
    } else {
        serde_json::to_writer(&mut out, &response)?;
        writeln!(out)?;
    }
    out.flush()?;
    Ok(ExitCode::SUCCESS)
}

/// The workbook of the command, with the locale that the configuration gives it: a `--locale` goes in
/// the query and wins over it.
fn book(cli: &Args) -> Result<Book> {
    let mut book = Book::from_file(&cli.file)?;
    if let Some(tag) = sheetrs_io::locale::configured(cli.config.as_deref())? {
        book = book.with_locale(tag);
    }
    Ok(book)
}

/// `--list-sheets` and `--list-names`: what the reader knows of the workbook, as JSON, with nothing
/// else added. The sheets come from the selective reader without reading a cell, and the names from
/// the reader of the defined names.
fn list(cli: &Args) -> Result<ExitCode> {
    let source = Book::from_file(&cli.file)?;
    let value = if cli.list_sheets {
        // A window on a sheet that does not exist: no sheet is opened, and the reader still lists them all.
        let selection = Selection::new(Properties::VALUE).window("\u{0}", Window::new(0, 0, 0, 0));
        let mut sink = Discard;
        let sheets = source.selection_into(&selection, &mut sink)?;
        serde_json::to_value(sheets.into_iter().map(|s| s.name).collect::<Vec<_>>())?
    } else {
        let (names, _) = source.defined_names()?;
        let sorted: BTreeMap<_, _> = names.into_iter().collect();
        serde_json::to_value(sorted)?
    };
    let out = std::io::stdout();
    let mut out = std::io::BufWriter::new(out.lock());
    if cli.pretty {
        serde_json::to_writer_pretty(&mut out, &value)?;
    } else {
        serde_json::to_writer(&mut out, &value)?;
    }
    writeln!(out)?;
    out.flush()?;
    Ok(ExitCode::SUCCESS)
}

/// A receiver of cells that keeps none: the list of the sheets is all that is wanted.
struct Discard;

impl CellSink for Discard {
    fn cell(&mut self, _sheet: u16, _at: (u32, u32), _cell: Cell) {}
}

#![warn(dead_code_pub_in_binary)]

use anyhow::{Context, Result, bail};
use clap::{Parser, ValueEnum};
use sheetrs_io::reader::Source;
use sheetrs_protocol::{Find, Mode, Property, Query, Reference, Scope, SearchIn};
use std::io::Write;
use std::path::PathBuf;

#[derive(Clone, Copy, ValueEnum)]
enum Prop {
    Value,
    Formula,
    #[value(alias = "number_format")]
    NumberFormat,
    Display,
}

#[derive(Clone, Copy, ValueEnum)]
enum Where {
    Value,
    Formula,
}

#[derive(Parser)]
#[command(name = "sheetquery")]
#[command(
    about = "Read the cells of an XLSX or ODS workbook by range, sheet or workbook, with the properties asked for, and find the cells that match a search",
    long_about = None
)]
#[command(version)]
#[allow(clippy::struct_excessive_bools)] // the flags of a command line
struct Cli {
    /// Path to the XLSX/ODS file
    #[arg(value_name = "FILE")]
    file: PathBuf,

    /// A range to read (`Sheet1!A1:C9`); without a sheet, the first one. Repeatable
    #[arg(long, value_name = "RANGE")]
    range: Vec<String>,

    /// A sheet to read whole. Repeatable
    #[arg(long, value_name = "SHEET")]
    sheet: Vec<String>,

    /// Read every sheet (the default when no range or sheet is given)
    #[arg(long)]
    workbook: bool,

    /// The properties of each cell; the value when omitted. Comma separated or repeated
    #[arg(
        short,
        long,
        value_enum,
        value_delimiter = ',',
        value_name = "PROPERTY"
    )]
    property: Vec<Prop>,

    /// Keep only the cells where this text is found (as `LibreOffice`'s Find & Replace finds it)
    #[arg(long, value_name = "TEXT")]
    find: Option<String>,

    /// The search is a regular expression (ICU, the subset that is reproduced)
    #[arg(long, conflicts_with = "wildcards")]
    regex: bool,

    /// The search has wildcards (`*`, `?`, `\` to escape)
    #[arg(long)]
    wildcards: bool,

    /// Distinguish upper and lower case in the search
    #[arg(long)]
    match_case: bool,

    /// The search must match the whole content of the cell
    #[arg(long)]
    entire_cells: bool,

    /// What the search looks at
    #[arg(long = "in", value_enum, default_value = "value")]
    search_in: Where,

    /// A search by value also looks at the result of the formulas
    #[arg(long)]
    include_formulas: bool,

    /// Locale of `display` and of the dates a search sees (default `en-US`)
    #[arg(long, value_name = "TAG")]
    locale: Option<String>,

    /// The whole query as JSON (a text, or `@file`); replaces the options above
    #[arg(long, value_name = "JSON")]
    query: Option<String>,

    /// One JSON object per cell, one per line
    #[arg(long)]
    jsonl: bool,

    /// Indent the JSON
    #[arg(long)]
    pretty: bool,
}

fn build(cli: &Cli) -> Result<Query> {
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
            Reference::parse(r).map_err(|e| anyhow::anyhow!("{e}"))?,
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
        },
        include_formulas: cli.include_formulas,
    });
    if find.is_none()
        && (cli.regex
            || cli.wildcards
            || cli.match_case
            || cli.entire_cells
            || cli.include_formulas)
    {
        bail!("the search options need --find");
    }
    Ok(Query {
        scope,
        properties,
        find,
        locale: cli.locale.clone(),
    })
}

fn main() -> Result<()> {
    let cli = Cli::parse();
    let query = build(&cli)?;
    let source = Source::from_file(&cli.file)?;
    let response = sheetrs_query::run(&source, &query)?;
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
    Ok(())
}

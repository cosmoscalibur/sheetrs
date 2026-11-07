//! Answering a query: select, read, filter, say.

use crate::error::{Error, Result};
use crate::matcher::Matcher;
use crate::text::TextContext;
use sheetrs_formula::addr::{a1, same_name};
use sheetrs_io::reader::{
    Cell, CellSink, CellValue, Properties, Selection, SheetInfo, Source, Window,
};
use sheetrs_protocol::{CellRecord, Find, Property, Query, QueryResponse, Scope};

/// The cells the reader hands over, with what is needed to filter them.
struct Collect<'a> {
    cells: Vec<(u16, u32, u32, Cell)>,
    /// The search, and the names of the sheets by index so far.
    find: Option<(&'a Find, Matcher)>,
    names: Vec<String>,
    text: &'a TextContext,
}

impl CellSink for Collect<'_> {
    fn cell(&mut self, sheet: u16, at: (u32, u32), cell: Cell) {
        if matches!(cell.value, CellValue::Empty) && cell.formula.is_none() {
            return;
        }
        if let Some((find, matcher)) = &self.find {
            let name = self
                .names
                .get(usize::from(sheet))
                .map_or("", String::as_str);
            let hit = self
                .text
                .searched(&cell, name, find.search_in, find.include_formulas)
                .is_some_and(|t| matcher.is_match(&t));
            if !hit {
                return;
            }
        }
        self.cells.push((sheet, at.0, at.1, cell));
    }
}

/// The name of the first sheet of the workbook.
fn first_sheet(source: &Source) -> Result<String> {
    struct Nothing;
    impl CellSink for Nothing {
        fn cell(&mut self, _: u16, _: (u32, u32), _: Cell) {}
    }
    let infos = source.read_selection_into(&Selection::new(Properties::VALUE), &mut Nothing)?;
    infos
        .into_iter()
        .next()
        .map(|i| i.name)
        .ok_or_else(|| Error::Request("the workbook has no sheets".into()))
}

fn window(r: &sheetrs_protocol::Reference) -> Window {
    Window::new(r.area.r0, r.area.c0, r.area.r1, r.area.c1)
}

/// Answers `query` over the workbook in `source`.
///
/// The cells come sheet by sheet in the order of the workbook, each sheet by rows; empty cells are
/// not returned.
///
/// # Errors
///
/// When the workbook cannot be read, the query names a sheet it does not have, the locale is
/// unknown, or the search is one the matcher does not reproduce.
pub fn run(source: &Source, query: &Query) -> Result<QueryResponse> {
    let wants = |p: Property| query.properties.contains(&p);
    let searching = query.find.is_some();
    let properties = Properties {
        // Whether a cell is empty is told by its value: it is read even when it is not asked for.
        value: true,
        formula: wants(Property::Formula) || searching,
        number_format: wants(Property::NumberFormat) || wants(Property::Display) || searching,
    };

    let mut selection = Selection::new(properties);
    let mut named: Vec<String> = Vec::new();
    for scope in &query.scope {
        match scope {
            Scope::Workbook => selection = selection.window_everywhere(Window::ALL),
            Scope::Sheet(name) => {
                named.push(name.clone());
                selection = selection.window(name.clone(), Window::ALL);
            }
            Scope::Range(reference) => {
                let sheet = match &reference.sheet {
                    Some(name) => name.clone(),
                    None => first_sheet(source)?,
                };
                named.push(sheet.clone());
                selection = selection.window(sheet, window(reference));
            }
        }
    }

    let needs_text = wants(Property::Display) || searching;
    let text = if needs_text {
        TextContext::new(query.locale.as_deref(), source.uses_1904_dates()?)?
    } else {
        TextContext::new(None, false)?
    };
    let matcher = query
        .find
        .as_ref()
        .map(|f| Matcher::new(f).map(|m| (f, m)))
        .transpose()?;

    // A search in formulas needs the name of the sheet of each cell (the formula says `$Sheet.A1`
    // unless the sheet is its own): the names are read first, and only then (an ODS is read through
    // to list them).
    let names: Vec<String> = if query
        .find
        .as_ref()
        .is_some_and(|f| f.search_in == sheetrs_protocol::SearchIn::Formula)
    {
        sheet_names(source)?
    } else {
        Vec::new()
    };

    let mut sink = Collect {
        cells: Vec::new(),
        find: matcher,
        names,
        text: &text,
    };
    let infos = source.read_selection_into(&selection, &mut sink)?;
    for name in &named {
        if !infos.iter().any(|i| same_name(&i.name, name)) {
            return Err(Error::Request(format!(
                "the workbook has no sheet `{name}`"
            )));
        }
    }

    let mut cells = sink.cells;
    cells.sort_by_key(|(sheet, row, col, _)| (*sheet, *row, *col));
    let records = cells
        .into_iter()
        .map(|(sheet, row, col, cell)| record(&infos, &text, query, sheet, row, col, &cell))
        .collect();
    Ok(QueryResponse { cells: records })
}

fn sheet_names(source: &Source) -> Result<Vec<String>> {
    let infos: Vec<SheetInfo> = {
        struct Nothing;
        impl CellSink for Nothing {
            fn cell(&mut self, _: u16, _: (u32, u32), _: Cell) {}
        }
        source.read_selection_into(&Selection::new(Properties::VALUE), &mut Nothing)?
    };
    let mut names = vec![
        String::new();
        infos
            .iter()
            .map(|i| usize::from(i.index) + 1)
            .max()
            .unwrap_or(0)
    ];
    for i in infos {
        names[usize::from(i.index)] = i.name;
    }
    Ok(names)
}

fn json(value: &CellValue) -> serde_json::Value {
    match value {
        CellValue::Empty => serde_json::Value::Null,
        CellValue::Number(x) => serde_json::Number::from_f64(*x).map_or_else(
            || serde_json::Value::String(x.to_string()),
            serde_json::Value::Number,
        ),
        CellValue::Text(t) => serde_json::Value::String(t.to_string()),
        CellValue::Boolean(b) => serde_json::Value::Bool(*b),
        CellValue::Error(e) => serde_json::Value::String(e.to_string()),
    }
}

fn record(
    infos: &[SheetInfo],
    text: &TextContext,
    query: &Query,
    sheet: u16,
    row: u32,
    col: u32,
    cell: &Cell,
) -> CellRecord {
    let wants = |p: Property| query.properties.contains(&p);
    CellRecord {
        sheet: infos
            .iter()
            .find(|i| i.index == sheet)
            .map(|i| i.name.clone())
            .unwrap_or_default(),
        cell: a1(row, col),
        row: row + 1,
        col: col + 1,
        value: wants(Property::Value).then(|| json(&cell.value)),
        formula: if wants(Property::Formula) {
            cell.formula.as_deref().map(str::to_string)
        } else {
            None
        },
        // What `LibreOffice` reports as the format of the cell: `General` when the file says none (an
        // ODS leaves it out, an XLSX writes it) and `BOOLEAN` for a boolean.
        number_format: if wants(Property::NumberFormat) {
            Some(match (&cell.value, cell.num_fmt.as_deref()) {
                (CellValue::Boolean(_), None) => "BOOLEAN".to_string(),
                (CellValue::Boolean(_), Some(f)) if f.eq_ignore_ascii_case("general") => {
                    "BOOLEAN".to_string()
                }
                (_, Some(f)) => f.to_string(),
                (_, None) => "General".to_string(),
            })
        } else {
            None
        },
        display: wants(Property::Display).then(|| text.display(cell)),
    }
}

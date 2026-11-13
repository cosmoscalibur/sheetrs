//! Answering a query: select, read, filter, say.

use crate::blocks::{Block, FindOptions, RangeOptions};
use crate::error::{Error, Result};
use crate::matcher::Matcher;
use crate::text::TextContext;
use sheetrs_formula::DateSystem;
use sheetrs_formula::addr::same_name;
use sheetrs_io::Book;
use sheetrs_io::reader::{Cell, CellSink, CellValue, Properties, Selection, SheetInfo, Window};
use sheetrs_protocol::{CellRecord, Find, Property, Query, QueryResponse, Scope};
use std::sync::Arc;

/// The cells the reader hands over, with what is needed to filter them.
struct Collect<'a> {
    cells: Vec<(u16, u32, u32, Cell)>,
    /// The notes that are asked for or searched, by cell: they come before their cells.
    notes: std::collections::HashMap<(u16, u32, u32), String>,
    /// Whether the search looks in the notes (and keeps only the cells whose note matches).
    in_notes: bool,
    /// The search, and the names of the sheets by index so far.
    find: Option<(&'a Find, Matcher)>,
    names: Vec<String>,
    text: &'a TextContext,
    /// Formula cells kept for a search of their result: its text depends on the format the cell
    /// takes from it, known only after the formats of the results are worked out.
    deferred: std::collections::HashSet<(u16, u32, u32)>,
    /// The rows an autofilter hid, by sheet: left out of a search unless it asks for them.
    filtered: Vec<(u16, u32, u32)>,
}

impl CellSink for Collect<'_> {
    fn filtered_rows(&mut self, sheet: u16, first: u32, last: u32) {
        self.filtered.push((sheet, first, last));
    }

    fn comment(&mut self, sheet: u16, at: (u32, u32), text: &str) {
        if self.in_notes
            && !self
                .find
                .as_ref()
                .is_some_and(|(_, matcher)| matcher.is_match(text))
        {
            return;
        }
        self.notes.insert((sheet, at.0, at.1), text.to_string());
    }

    fn cell(&mut self, sheet: u16, at: (u32, u32), cell: Cell) {
        if matches!(cell.value, CellValue::Empty) && cell.formula.is_none() {
            return;
        }
        if self.in_notes {
            if self.notes.contains_key(&(sheet, at.0, at.1)) {
                self.cells.push((sheet, at.0, at.1, cell));
            }
            return;
        }
        if let Some((find, matcher)) = &self.find {
            if cell.formula.is_some()
                && find.search_in == sheetrs_protocol::SearchIn::Value
                && find.include_formulas
            {
                self.deferred.insert((sheet, at.0, at.1));
                self.cells.push((sheet, at.0, at.1, cell));
                return;
            }
            let name = self
                .names
                .get(usize::from(sheet))
                .map_or("", String::as_str);
            let hit = self
                .text
                .searched_as(
                    &cell,
                    name,
                    find.search_in,
                    find.include_formulas,
                    find.formatted,
                )
                .is_some_and(|t| matcher.is_match(&t));
            if !hit {
                return;
            }
        }
        self.cells.push((sheet, at.0, at.1, cell));
    }
}

/// The name of the first sheet of the workbook.
fn first_sheet(source: &Book) -> Result<String> {
    struct Nothing;
    impl CellSink for Nothing {
        fn cell(&mut self, _: u16, _: (u32, u32), _: Cell) {}
    }
    let infos = source.selection_into(&Selection::new(Properties::VALUE), &mut Nothing)?;
    infos
        .into_iter()
        .next()
        .map(|i| i.name)
        .ok_or_else(|| Error::Request("the workbook has no sheets".into()))
}

fn window(r: &sheetrs_protocol::Reference) -> Window {
    Window::new(r.area.r0, r.area.c0, r.area.r1, r.area.c1)
}

/// What a [`Book`] answers: the cells a [`Query`] selects.
pub trait BookQuery {
    /// Answers `query` over the workbook.
    ///
    /// The cells come sheet by sheet in the order of the workbook, each sheet by rows; empty cells
    /// are not returned.
    ///
    /// # Errors
    ///
    /// When the workbook cannot be read, the query names a sheet it does not have, the locale is
    /// unknown, or the search is one the matcher does not reproduce.
    fn query(&self, query: &Query) -> Result<QueryResponse>;

    /// The cells of one or several ranges, read in a single pass over the file, as one [`Block`] per
    /// range in the order asked (`sheetrs query --range`, `Book.range` of Python).
    ///
    /// A range is `Data!A1:C10`, a range open at one end (`Data!A:C`, `Data!3:10`) or a defined name
    /// that is a range; a range with no sheet is on the first sheet. Only the properties asked for
    /// are read.
    ///
    /// ```no_run
    /// use sheetrs_io::Book;
    /// use sheetrs_query::{BookQuery, RangeOptions};
    ///
    /// # fn main() -> sheetrs_query::Result<()> {
    /// let book = Book::from_file("book.xlsx")?;
    /// let blocks = book.range(["Data!A1:C10"], &RangeOptions::default())?;
    /// println!("{:?}", blocks[0].values);
    /// # Ok(()) }
    /// ```
    ///
    /// # Errors
    ///
    /// When a range is neither a range nor a defined name, names a sheet the workbook does not
    /// have, or the workbook cannot be read.
    fn range<I, S>(&self, ranges: I, options: &RangeOptions) -> Result<Vec<Block>>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>;

    /// The cells that match a search, as Find & Replace of `LibreOffice` does (`sheetrs query
    /// --find`, `Book.find` of Python): the cells with the properties asked for, by rows or by
    /// columns.
    ///
    /// ```no_run
    /// use sheetrs_io::Book;
    /// use sheetrs_protocol::Find;
    /// use sheetrs_query::{BookQuery, FindOptions};
    ///
    /// # fn main() -> sheetrs_query::Result<()> {
    /// let book = Book::from_file("book.xlsx")?;
    /// for cell in book.find(&Find::literal("total"), &FindOptions::default())? {
    ///     println!("{}!{}", cell.sheet, cell.cell());
    /// }
    /// # Ok(()) }
    /// ```
    ///
    /// # Errors
    ///
    /// As [`query`](Self::query).
    fn find(&self, find: &Find, options: &FindOptions) -> Result<Vec<CellRecord>>;
}

impl BookQuery for Book {
    fn query(&self, query: &Query) -> Result<QueryResponse> {
        run(self, query)
    }

    fn range<I, S>(&self, ranges: I, options: &RangeOptions) -> Result<Vec<Block>>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        crate::blocks::range(self, ranges, options)
    }

    fn find(&self, find: &Find, options: &FindOptions) -> Result<Vec<CellRecord>> {
        crate::blocks::find(self, find, options)
    }
}

/// The locale of a query: the one it states, else the one given to the book, else the one of the
/// system (for an XLSX and an ODS alike). A system locale that the engine has no locale for is
/// `en-US`, as it is for `LibreOffice`.
fn locale_of(source: &Book, query: &Query) -> Option<String> {
    if let Some(tag) = &query.locale {
        return Some(tag.clone());
    }
    if let Some(tag) = source.locale() {
        return Some(tag.to_string());
    }
    source
        .default_locale()
        .map(|tag| tag.to_string())
        .filter(|tag| sheetrs_engine::options::Locale::by_tag(tag).is_some())
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
pub(crate) fn run(source: &Book, query: &Query) -> Result<QueryResponse> {
    let wants = |p: Property| query.properties.contains(&p);
    let searching = query.find.is_some();
    // Only the value is given as a date, and only when it is asked for.
    let as_dates = query.dates && wants(Property::Value);
    let in_notes = query
        .find
        .as_ref()
        .is_some_and(|f| f.search_in == sheetrs_protocol::SearchIn::Comment);
    let properties = Properties {
        // Whether a cell is empty is told by its value: it is read even when it is not asked for.
        value: true,
        formula: wants(Property::Formula) || searching,
        // A date is told from a number by its format.
        number_format: wants(Property::NumberFormat)
            || wants(Property::Display)
            || searching
            || as_dates,
        comment: wants(Property::Comment) || in_notes,
    };

    let mut selection = Selection::new(properties);
    // A search leaves out the rows an autofilter hid, unless it asks for them.
    let leaves_out_filtered = query.find.as_ref().is_some_and(|f| !f.filtered);
    if leaves_out_filtered {
        selection = selection.with_filtered_rows();
    }
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

    let needs_text = wants(Property::Display) || searching || as_dates;
    let text = if needs_text {
        TextContext::new(locale_of(source, query).as_deref(), source.date_system()?)?
            .for_format(source.format())
            .for_book(source)
    } else {
        TextContext::new(None, DateSystem::default())?
            .for_format(source.format())
            .for_book(source)
    };
    let matcher = query
        .find
        .as_ref()
        .map(|f| Matcher::for_find(f).map(|m| (f, m)))
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
        notes: std::collections::HashMap::new(),
        in_notes,
        find: matcher,
        names,
        text: &text,
        deferred: std::collections::HashSet::new(),
        filtered: Vec::new(),
    };
    let infos = source.selection_into(&selection, &mut sink)?;
    for name in &named {
        if !infos.iter().any(|i| same_name(&i.name, name)) {
            return Err(Error::Request(format!(
                "the workbook has no sheet `{name}`"
            )));
        }
    }

    let mut cells = sink.cells;
    if leaves_out_filtered && !sink.filtered.is_empty() {
        let hidden = |sheet: u16, row: u32| {
            sink.filtered
                .iter()
                .any(|&(s, first, last)| s == sheet && (first..=last).contains(&row))
        };
        cells.retain(|(sheet, row, _, _)| !hidden(*sheet, *row));
        sink.notes.retain(|&(sheet, row, _), _| !hidden(sheet, row));
    }
    // A note on a cell that has nothing else is a cell too, when the notes are what is asked.
    if in_notes || (wants(Property::Comment) && !searching) {
        let seen: std::collections::HashSet<(u16, u32, u32)> =
            cells.iter().map(|(s, r, c, _)| (*s, *r, *c)).collect();
        for &(sheet, row, col) in sink.notes.keys() {
            if !seen.contains(&(sheet, row, col)) {
                cells.push((sheet, row, col, Cell::default()));
            }
        }
    }
    cells.sort_by_key(|(sheet, row, col, _)| (*sheet, *row, *col));
    if wants(Property::NumberFormat)
        || wants(Property::Display)
        || as_dates
        || !sink.deferred.is_empty()
    {
        inherit_formats(source, &infos, &text, &mut cells);
    }
    if let Some((find, matcher)) = &sink.find
        && !sink.deferred.is_empty()
    {
        cells.retain(|(sheet, row, col, cell)| {
            if !sink.deferred.contains(&(*sheet, *row, *col)) {
                return true;
            }
            let name = infos
                .iter()
                .find(|i| i.index == *sheet)
                .map_or("", |i| i.name.as_str());
            text.searched_as(
                cell,
                name,
                find.search_in,
                find.include_formulas,
                find.formatted,
            )
            .is_some_and(|t| matcher.is_match(&t))
        });
    }
    // One copy of each sheet name, shared by the cells of the sheet.
    let sheets: Vec<(u16, Arc<str>)> = infos
        .iter()
        .map(|i| (i.index, Arc::from(i.name.as_str())))
        .collect();
    let records = cells
        .into_iter()
        .map(|(sheet, row, col, cell)| {
            let note = sink.notes.get(&(sheet, row, col)).map(String::as_str);
            record(&sheets, &text, query, (sheet, row, col), &cell, note)
        })
        .collect();
    Ok(QueryResponse { cells: records })
}

/// Gives the formula cells that have no format the one `LibreOffice` takes from their result.
fn inherit_formats(
    source: &Book,
    infos: &[SheetInfo],
    text: &TextContext,
    cells: &mut [(u16, u32, u32, Cell)],
) {
    if !cells.iter().any(|(_, _, _, c)| c.formula.is_some()) {
        return;
    }
    let found: Vec<(usize, String)> = {
        let mut inference = crate::inherit::Inference::new(source, infos, cells, text);
        cells
            .iter()
            .enumerate()
            .filter_map(|(i, (s, r, c, cell))| {
                inference.format_of((*s, *r, *c), cell).map(|f| (i, f))
            })
            .collect()
    };
    for (i, format) in found {
        cells[i].3.num_fmt = Some(format.into());
    }
}

fn sheet_names(source: &Book) -> Result<Vec<String>> {
    let infos: Vec<SheetInfo> = {
        struct Nothing;
        impl CellSink for Nothing {
            fn cell(&mut self, _: u16, _: (u32, u32), _: Cell) {}
        }
        source.selection_into(&Selection::new(Properties::VALUE), &mut Nothing)?
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
        // An error is not a text that starts with `#`: it is told apart, as the answer of `sheetc calculate`
        // tells it (`{"error": "#DIV/0!"}`).
        CellValue::Error(e) => serde_json::json!({"error": e.to_string()}),
    }
}

fn record(
    sheets: &[(u16, Arc<str>)],
    text: &TextContext,
    query: &Query,
    (sheet, row, col): (u16, u32, u32),
    cell: &Cell,
    note: Option<&str>,
) -> CellRecord {
    let wants = |p: Property| query.properties.contains(&p);
    CellRecord {
        sheet: sheets
            .iter()
            .find(|(index, _)| *index == sheet)
            .map_or_else(|| Arc::from(""), |(_, name)| Arc::clone(name)),
        row: row + 1,
        col: col + 1,
        value: wants(Property::Value).then(|| match &cell.value {
            CellValue::Number(x) if query.dates => {
                crate::dates::typed(*x, cell.num_fmt.as_deref(), text.date_system())
                    .map_or_else(|| json(&cell.value), |d| d.to_json())
            }
            _ => json(&cell.value),
        }),
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
        comment: if wants(Property::Comment) {
            note.map(str::to_string)
        } else {
            None
        },
    }
}

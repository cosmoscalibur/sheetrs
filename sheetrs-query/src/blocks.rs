//! `range` and `find`: the reads of a workbook that the command line (`sheetrs query --range`,
//! `--find`), Python (`Book.range`, `Book.find`) and WebAssembly give, as methods of
//! [`BookQuery`]. They are the [`Query`] of the protocol asked in
//! the shape each is used in: the cells of ranges as blocks, the cells that match a search as a
//! list.

use crate::error::{Error, Result};
use crate::run::BookQuery;
use serde_json::Value as Json;
use sheetrs_io::Book;
use sheetrs_protocol::{CellRecord, Find, Property, Query, Reference, Scope};
use std::collections::HashMap;

const LAST_ROW: u32 = (1 << 20) - 1;
const LAST_COLUMN: u32 = (1 << 14) - 1;

/// Whether a block, or a list of matches, is given by rows or by columns.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Major {
    /// A list of rows, each a list of cells (the default).
    #[default]
    Rows,
    /// A list of columns, each a list of cells.
    Columns,
}

/// What [`BookQuery::range`] is asked besides the ranges.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RangeOptions {
    /// The properties of each cell: the value when omitted.
    pub properties: Vec<Property>,
    /// Rows (the default) or columns.
    pub major: Major,
    /// Give the dates and times as typed values (`false`, the default: the serial numbers).
    pub dates: bool,
}

impl Default for RangeOptions {
    fn default() -> Self {
        Self {
            properties: vec![Property::Value],
            major: Major::Rows,
            dates: false,
        }
    }
}

/// What [`BookQuery::find`] is asked besides the search.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FindOptions {
    /// A sheet, a range or a list of them: the whole workbook when empty.
    pub scope: Vec<Scope>,
    /// The properties of each match: the value when omitted.
    pub properties: Vec<Property>,
    /// The matches by rows (the default) or by columns, sheet by sheet.
    pub by: Major,
    /// Give the dates and times as typed values (`false`, the default: the serial numbers).
    pub dates: bool,
}

impl Default for FindOptions {
    fn default() -> Self {
        Self {
            scope: Vec::new(),
            properties: vec![Property::Value],
            by: Major::Rows,
            dates: false,
        }
    }
}

/// The cells of one range: one grid per property that was asked for, as lists of rows (or of
/// columns, with [`Major::Columns`]).
///
/// A bounded range (`A1:C10`) is as big as the range, with `None` where a cell is empty; a range
/// open at one end (`A:C`, `3:10`) is cut where its last non-empty cell is.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Block {
    /// The range as it was asked (a defined name stays a name).
    pub range: String,
    /// How the grids are laid out.
    pub major: Major,
    /// The values, when asked for.
    pub values: Option<Vec<Vec<Option<Json>>>>,
    /// The formulas, when asked for.
    pub formulas: Option<Vec<Vec<Option<String>>>>,
    /// The number formats, when asked for.
    pub number_formats: Option<Vec<Vec<Option<String>>>>,
    /// The texts the cells show, when asked for.
    pub displays: Option<Vec<Vec<Option<String>>>>,
    /// The notes, when asked for.
    pub comments: Option<Vec<Vec<Option<String>>>>,
}

struct Wanted {
    key: String,
    reference: Reference,
}

/// A range as written (`Data!A1:C10`) or a defined name that is one.
fn resolve(book: &Book, text: &str) -> Result<Wanted> {
    let reference = if let Ok(reference) = text.parse::<Reference>() {
        reference
    } else {
        let (names, _) = book.defined_names()?;
        let refers_to = names
            .iter()
            .find(|(name, _)| name.eq_ignore_ascii_case(text.trim()))
            .map(|(_, refers_to)| refers_to.clone())
            .ok_or_else(|| {
                Error::Request(format!("`{text}` is neither a range nor a defined name"))
            })?;
        refers_to
            .trim_start_matches('=')
            .parse::<Reference>()
            .map_err(|_| {
                Error::Request(format!(
                    "the name `{text}` refers to `{refers_to}`, which is not a single range"
                ))
            })?
    };
    Ok(Wanted {
        key: text.to_owned(),
        reference,
    })
}

fn grid<T: Clone>(rows: Vec<Vec<T>>, major: Major) -> Vec<Vec<T>> {
    if major == Major::Columns && !rows.is_empty() {
        let width = rows.iter().map(Vec::len).max().unwrap_or(0);
        let mut out: Vec<Vec<T>> = (0..width).map(|_| Vec::new()).collect();
        for row in rows {
            for (c, cell) in row.into_iter().enumerate() {
                out[c].push(cell);
            }
        }
        out
    } else {
        rows
    }
}

/// The block of one requested range from the cells of the answer.
fn block_of(
    wanted: &Wanted,
    cells: &HashMap<(u32, u32), &CellRecord>,
    options: &RangeOptions,
) -> Block {
    let area = &wanted.reference.area;
    let (r0, c0) = (area.r0, area.c0);
    let wants = |p: Property| options.properties.contains(&p);
    let mut block = Block {
        range: wanted.key.clone(),
        major: options.major,
        values: wants(Property::Value).then(Vec::new),
        formulas: wants(Property::Formula).then(Vec::new),
        number_formats: wants(Property::NumberFormat).then(Vec::new),
        displays: wants(Property::Display).then(Vec::new),
        comments: wants(Property::Comment).then(Vec::new),
    };
    // The cells of an open range are cut where the last one is.
    let last = |pick: fn(&CellRecord) -> u32, from: u32| {
        cells
            .values()
            .map(|c| pick(c) - 1)
            .filter(|&n| n >= from)
            .max()
    };
    let r1 = if area.r1 >= LAST_ROW {
        last(|c| c.row, r0)
    } else {
        Some(area.r1)
    };
    let c1 = if area.c1 >= LAST_COLUMN {
        last(|c| c.col, c0)
    } else {
        Some(area.c1)
    };
    // A range open at one end with no cell in it has no rows (or no columns) to give.
    let (Some(r1), Some(c1)) = (r1, c1) else {
        return block;
    };
    let (mut v, mut f, mut n, mut d, mut m) = (vec![], vec![], vec![], vec![], vec![]);
    for row in r0..=r1 {
        let (mut vr, mut fr, mut nr, mut dr, mut mr) = (vec![], vec![], vec![], vec![], vec![]);
        for col in c0..=c1 {
            let record = cells.get(&(row + 1, col + 1)).copied();
            vr.push(record.and_then(|r| r.value.clone()));
            fr.push(record.and_then(|r| r.formula.clone()));
            nr.push(record.and_then(|r| r.number_format.clone()));
            dr.push(record.and_then(|r| r.display.clone()));
            mr.push(record.and_then(|r| r.comment.clone()));
        }
        v.push(vr);
        f.push(fr);
        n.push(nr);
        d.push(dr);
        m.push(mr);
    }
    let major = options.major;
    block.values = block.values.map(|_| grid(v, major));
    block.formulas = block.formulas.map(|_| grid(f, major));
    block.number_formats = block.number_formats.map(|_| grid(n, major));
    block.displays = block.displays.map(|_| grid(d, major));
    block.comments = block.comments.map(|_| grid(m, major));
    block
}

pub(crate) fn range<I, S>(book: &Book, ranges: I, options: &RangeOptions) -> Result<Vec<Block>>
where
    I: IntoIterator<Item = S>,
    S: AsRef<str>,
{
    let wanted: Vec<Wanted> = ranges
        .into_iter()
        .map(|text| resolve(book, text.as_ref()))
        .collect::<Result<_>>()?;
    let query = Query {
        scope: wanted
            .iter()
            .map(|w| Scope::Range(w.reference.clone()))
            .collect(),
        properties: options.properties.clone(),
        find: None,
        locale: None,
        dates: options.dates,
    };
    let response = book.query(&query)?;
    // The sheet a range with no sheet is on: the first of the workbook.
    let first = if wanted.iter().any(|w| w.reference.sheet.is_none()) {
        book.sheet_names()?.into_iter().next().unwrap_or_default()
    } else {
        String::new()
    };
    let mut by_sheet: HashMap<String, HashMap<(u32, u32), &CellRecord>> = HashMap::new();
    for cell in &response.cells {
        by_sheet
            .entry(cell.sheet.to_lowercase())
            .or_default()
            .insert((cell.row, cell.col), cell);
    }
    let empty = HashMap::new();
    Ok(wanted
        .iter()
        .map(|w| {
            let sheet = w.reference.sheet.clone().unwrap_or_else(|| first.clone());
            let cells = by_sheet.get(&sheet.to_lowercase()).unwrap_or(&empty);
            // The cells of this range only: the ones of the answer that are inside it.
            let area = &w.reference.area;
            let inside: HashMap<(u32, u32), &CellRecord> = cells
                .iter()
                .filter(|((row, col), _)| {
                    (area.r0..=area.r1).contains(&(row - 1))
                        && (area.c0..=area.c1).contains(&(col - 1))
                })
                .map(|(k, v)| (*k, *v))
                .collect();
            block_of(w, &inside, options)
        })
        .collect())
}

pub(crate) fn find(book: &Book, find: &Find, options: &FindOptions) -> Result<Vec<CellRecord>> {
    let query = Query {
        scope: if options.scope.is_empty() {
            vec![Scope::Workbook]
        } else {
            options.scope.clone()
        },
        properties: options.properties.clone(),
        find: Some(find.clone()),
        locale: None,
        dates: options.dates,
    };
    let mut cells = book.query(&query)?.cells;
    if options.by == Major::Columns {
        // The answer is sheet by sheet; the sheets stay in the order they came in.
        let mut order: Vec<std::sync::Arc<str>> = Vec::new();
        for cell in &cells {
            if !order.contains(&cell.sheet) {
                order.push(std::sync::Arc::clone(&cell.sheet));
            }
        }
        cells.sort_by_key(|c| {
            (
                order.iter().position(|s| *s == c.sheet).unwrap_or(0),
                c.col,
                c.row,
            )
        });
    }
    Ok(cells)
}

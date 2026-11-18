//! The structure of `content.xml`: tables, rows, cells and the paragraphs of their text.

use std::collections::{HashMap, HashSet};
use std::io::BufRead;

use sheetrs_core::{
    CellPos, DateSystem, FormulaRef, MAX_COLS, MAX_ROWS, SheetBuilder, ValueRef, Visibility,
    Workbook,
};
use sheetrs_io::{DateRules, integer_prefix, iso_day};

use crate::Error;
use crate::tokens::{Attrs, Tag, Token, Tokens};
use crate::value::{Scalar, scalar};

/// The most bytes the text of one paragraph may hold when it holds `text:s` elements.
const MAX_TEXT_BYTES: usize = 1 << 24;
/// The most elements one paragraph may nest, beyond which the reader stops.
const MAX_NESTING: usize = 256;
/// The most columns an array formula may span; a greater count drops the cell.
const MAX_MATRIX_COLUMNS: u32 = 32_767;
/// The most cells that a table may hold, repeated ones included.
pub(crate) const MAX_TABLE_CELLS: u64 = 1 << 26;

/// What the reader knows of the document while it reads the tables.
struct Document {
    system: DateSystem,
    /// Whether each table style shows its sheet, as its first definition says.
    shown_styles: HashMap<String, bool>,
}

/// The documents that formulas refer to, each once, in the order they were found.
#[derive(Default)]
struct Links {
    seen: HashSet<String>,
    found: Vec<String>,
}

impl Links {
    fn add(&mut self, target: String) {
        if self.seen.insert(target.clone()) {
            self.found.push(target);
        }
    }
}

/// How an array formula marks the cells it fills.
enum Matrix {
    /// The cell is not an array formula.
    Absent,
    /// The size is not one the sheet holds: the cell is dropped.
    Dropped,
    Area {
        columns: u32,
        rows: u32,
    },
}

/// A cell of the row being read.
struct CellData {
    scalar: Scalar,
    formula: Option<String>,
    stale: bool,
    matrix: Matrix,
    /// The columns the cell stands for.
    repeat: u64,
}

/// Adds to `links` the documents that the references of `formula` name: a quoted address
/// followed by `#`, outside the text of the formula.
fn external_documents(formula: &str, links: &mut Links) {
    if !formula.contains('\'') {
        return;
    }
    let mut rest = formula;
    while let Some(at) = rest.find(['\'', '"']) {
        let quote = rest.as_bytes().get(at).copied().unwrap_or(b'"');
        let quote_char = char::from(quote);
        let inside = rest.get(at + 1..).unwrap_or_default();
        // The quote closes at the first one that is not doubled.
        let mut end = 0;
        let mut doubled = false;
        let closed = loop {
            let Some(close) = inside.get(end..).and_then(|tail| tail.find(quote_char)) else {
                break None;
            };
            let close = end + close;
            if inside
                .get(close + 1..)
                .is_some_and(|tail| tail.starts_with(quote_char))
            {
                doubled = true;
                end = close + 2;
            } else {
                break Some(close);
            }
        };
        let Some(close) = closed else { return };
        let tail = inside.get(close + 1..).unwrap_or_default();
        if quote == b'\'' && tail.starts_with('#') {
            let name = inside.get(..close).unwrap_or_default();
            links.add(if doubled {
                name.replace("''", "'")
            } else {
                name.to_owned()
            });
        }
        rest = tail;
    }
}

pub(crate) fn parse<R: BufRead>(source: R, workbook: &mut Workbook) -> Result<(), Error> {
    let mut tokens = Tokens::new(source, crate::CONTENT_PART);
    let mut document = Document {
        system: DateSystem::Date1900,
        shown_styles: HashMap::new(),
    };
    let mut style: Option<String> = None;
    let mut links = Links::default();
    loop {
        match tokens.next()? {
            Token::Eof => {
                for target in &links.found {
                    workbook.add_external_link(target);
                }
                return Ok(());
            }
            Token::Start(Tag::Style, attrs) => {
                style = attrs
                    .name
                    .filter(|_| attrs.family.as_deref() == Some("table"));
            }
            Token::End(Tag::Style) => style = None,
            Token::Start(Tag::TableProperties, attrs)
            | Token::Empty(Tag::TableProperties, attrs) => {
                if let Some(name) = &style {
                    document
                        .shown_styles
                        .entry(name.clone())
                        .or_insert(attrs.display.as_deref() != Some("false"));
                }
            }
            Token::Start(Tag::NullDate, attrs) | Token::Empty(Tag::NullDate, attrs) => {
                document.system = null_date(attrs.null_date.as_deref());
                workbook.set_date_system(document.system);
            }
            Token::Start(Tag::Table, attrs) => {
                parse_table(&mut tokens, workbook, &document, &attrs, false, &mut links)?;
            }
            Token::Empty(Tag::Table, attrs) => {
                parse_table(&mut tokens, workbook, &document, &attrs, true, &mut links)?;
            }
            _ => {}
        }
    }
}

/// The date system of a `table:null-date`: the date it holds, or the 1900 system when the text
/// is not a date.
fn null_date(text: Option<&str>) -> DateSystem {
    let days_from_epoch =
        text.and_then(|text| iso_day(text, DateSystem::Custom { zero: 0 }, DateRules::ODS));
    match days_from_epoch {
        Some(-25_569) | None => DateSystem::Date1900,
        Some(-24_107) => DateSystem::Date1904,
        Some(zero) => DateSystem::Custom { zero },
    }
}

fn parse_table<R: BufRead>(
    tokens: &mut Tokens<R>,
    workbook: &mut Workbook,
    document: &Document,
    attrs: &Attrs,
    empty: bool,
    links: &mut Links,
) -> Result<(), Error> {
    let mut sheet = workbook.sheet_builder(attrs.name.clone().unwrap_or_default());
    if attrs
        .style_name
        .as_ref()
        .is_some_and(|style| document.shown_styles.get(style) == Some(&false))
    {
        sheet.set_visibility(Visibility::Hidden);
    }
    if !empty {
        parse_rows(tokens, document, &mut sheet, links)?;
    }
    sheet.finish();
    Ok(())
}

/// Reads the rows of the table just opened into `sheet`, up to the end of the table.
fn parse_rows<R: BufRead>(
    tokens: &mut Tokens<R>,
    document: &Document,
    sheet: &mut SheetBuilder<'_>,
    links: &mut Links,
) -> Result<(), Error> {
    let mut next_row: u64 = 0;
    let mut written: u64 = 0;
    loop {
        match tokens.next()? {
            Token::Start(Tag::Row, attrs) => {
                let repeat = repeated(attrs.rows_repeated.as_deref());
                let cells = parse_row(tokens, document, links)?;
                emit_row(sheet, next_row, repeat, &cells, &mut written)?;
                next_row += repeat;
            }
            Token::Empty(Tag::Row, attrs) => next_row += repeated(attrs.rows_repeated.as_deref()),
            Token::Start(Tag::RowContainer, _) | Token::End(Tag::RowContainer) => {}
            Token::Start(..) => tokens.skip()?,
            Token::End(Tag::Table) | Token::Eof => break,
            _ => {}
        }
    }
    Ok(())
}

/// The cells of the row just opened, each with the column it starts at.
fn parse_row<R: BufRead>(
    tokens: &mut Tokens<R>,
    document: &Document,
    links: &mut Links,
) -> Result<Vec<(u64, CellData)>, Error> {
    let mut cells = Vec::new();
    let mut column: u64 = 0;
    loop {
        let (attrs, paragraphs) = match tokens.next()? {
            Token::Start(Tag::Cell | Tag::CoveredCell, attrs) => {
                let paragraphs = cell_paragraphs(tokens)?;
                (attrs, paragraphs)
            }
            Token::Empty(Tag::Cell | Tag::CoveredCell, attrs) => (attrs, Vec::new()),
            Token::Start(..) => {
                tokens.skip()?;
                continue;
            }
            Token::End(Tag::Row) | Token::Eof => return Ok(cells),
            _ => continue,
        };
        let cell = cell_data(&attrs, &paragraphs, document);
        if let Some(text) = &cell.formula {
            external_documents(text, links);
        }
        let repeat = cell.repeat;
        if cell.scalar != Scalar::Empty || cell.formula.is_some() {
            cells.push((column, cell));
        }
        column += repeat;
    }
}

/// The paragraphs of the cell just opened, read up to its end.
fn cell_paragraphs<R: BufRead>(tokens: &mut Tokens<R>) -> Result<Vec<String>, Error> {
    let mut paragraphs = Vec::new();
    loop {
        match tokens.next()? {
            Token::Start(Tag::Paragraph, _) => paragraphs.push(paragraph(tokens)?),
            Token::Empty(Tag::Paragraph, _) => paragraphs.push(String::new()),
            Token::Start(..) => tokens.skip()?,
            Token::End(_) | Token::Eof => return Ok(paragraphs),
            _ => {}
        }
    }
}

/// Where in a paragraph the text being read is: `LibreOffice` takes tabs and line breaks only
/// from the paragraph itself, and drops the content of a span or link inside either.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Inline {
    Paragraph,
    Span,
    Link,
}

/// The text of the paragraph just opened: its characters, with each `text:s`, `text:tab` and
/// `text:line-break` as the characters they stand for.
fn paragraph<R: BufRead>(tokens: &mut Tokens<R>) -> Result<String, Error> {
    let mut text = String::new();
    inline(tokens, Inline::Paragraph, &mut text, 0)?;
    Ok(normalize_line_breaks(&text))
}

/// Appends the text of the element just opened, read up to its end, to `text`.
fn inline<R: BufRead>(
    tokens: &mut Tokens<R>,
    place: Inline,
    text: &mut String,
    depth: usize,
) -> Result<(), Error> {
    if depth > MAX_NESTING {
        return Err(Error::malformed(
            tokens.part(),
            "elements nested too deeply",
        ));
    }
    let start = text.len();
    loop {
        match tokens.next()? {
            Token::Text(content) => text.push_str(&content),
            Token::Start(Tag::Space, attrs) | Token::Empty(Tag::Space, attrs) => {
                let count = integer_prefix(attrs.count.as_deref().unwrap_or("1")).max(1);
                let count = usize::try_from(count).unwrap_or(1);
                if text.len().saturating_add(count) > MAX_TEXT_BYTES {
                    return Err(Error::malformed(
                        tokens.part(),
                        "the text of a paragraph is too long",
                    ));
                }
                text.extend(std::iter::repeat_n(' ', count));
            }
            Token::Start(Tag::Tab, _) | Token::Empty(Tag::Tab, _) if place == Inline::Paragraph => {
                text.push('\t');
            }
            Token::Start(Tag::LineBreak, _) | Token::Empty(Tag::LineBreak, _)
                if place == Inline::Paragraph =>
            {
                text.push('\n');
            }
            Token::Start(Tag::Span, _) if place == Inline::Paragraph => {
                inline(tokens, Inline::Span, text, depth + 1)?;
            }
            Token::Start(Tag::Link, _) if place != Inline::Link => {
                inline(tokens, Inline::Link, text, depth + 1)?;
            }
            Token::Start(Tag::Aside | Tag::Span | Tag::Link, _) => tokens.skip()?,
            Token::Start(..) => inline(tokens, place, text, depth + 1)?,
            Token::End(_) | Token::Eof => break,
            Token::Empty(..) => {}
        }
    }
    // A link without text still takes the place of one character.
    if place == Inline::Link && text.len() == start {
        text.push(' ');
    }
    Ok(())
}

/// A carriage return and a carriage return with a line feed are line feeds, as XML 1.0 says.
fn normalize_line_breaks(text: &str) -> String {
    if text.contains('\r') {
        text.replace("\r\n", "\n").replace('\r', "\n")
    } else {
        text.to_owned()
    }
}

/// The count of a `number-*-repeated` attribute: the integer it starts with, or one when that
/// is not a positive number.
fn repeated(attribute: Option<&str>) -> u64 {
    u64::try_from(integer_prefix(attribute.unwrap_or("1")).max(1)).unwrap_or(1)
}

fn cell_data(attrs: &Attrs, paragraphs: &[String], document: &Document) -> CellData {
    let (scalar, cached_text) = scalar(attrs, paragraphs, document.system);
    let formula = attrs.formula.clone();
    let matrix = match (&formula, &attrs.matrix_columns, &attrs.matrix_rows) {
        (Some(_), None, None) | (None, ..) => Matrix::Absent,
        (Some(_), columns, rows) => {
            let size = |text: &Option<String>| {
                u32::try_from(integer_prefix(text.as_deref().unwrap_or(""))).ok()
            };
            match (size(columns), size(rows)) {
                (Some(columns @ 1..=MAX_MATRIX_COLUMNS), Some(rows @ 1..)) => {
                    Matrix::Area { columns, rows }
                }
                _ => Matrix::Dropped,
            }
        }
    };
    CellData {
        // A formula cell is read once, whatever the repeat count says.
        repeat: if formula.is_some() {
            1
        } else {
            repeated(attrs.columns_repeated.as_deref())
        },
        // Only a text cached with its cell is kept; every other result is recalculated.
        stale: formula.is_some() && !cached_text,
        scalar,
        formula,
        matrix,
    }
}

/// How many cells `cell`, which starts at column `start`, stands for in one row.
fn cell_count(start: u64, cell: &CellData) -> u64 {
    let columns = u64::from(MAX_COLS).saturating_sub(start);
    match &cell.matrix {
        Matrix::Dropped => 0,
        Matrix::Area {
            columns: width,
            rows,
        } => u64::from(*width).min(columns) * u64::from((*rows).min(MAX_ROWS)),
        Matrix::Absent => cell.repeat.min(columns),
    }
}

/// Writes the cells of a row, repeated `repeat` times, into `sheet`.
fn emit_row(
    sheet: &mut SheetBuilder<'_>,
    first_row: u64,
    repeat: u64,
    cells: &[(u64, CellData)],
    written: &mut u64,
) -> Result<(), Error> {
    if cells.is_empty() {
        return Ok(());
    }
    // The cells are counted before any is written, so that a file that stands for far more cells
    // than it has bytes is refused at once.
    let rows = repeat.min(u64::from(MAX_ROWS).saturating_sub(first_row));
    let per_row: u64 = cells
        .iter()
        .map(|(start, cell)| cell_count(*start, cell))
        .sum();
    *written = written.saturating_add(rows.saturating_mul(per_row));
    if *written > MAX_TABLE_CELLS {
        return Err(Error::TooManyCells(MAX_TABLE_CELLS));
    }
    for row in first_row..first_row.saturating_add(repeat) {
        let Ok(row) = u32::try_from(row) else { break };
        if row >= MAX_ROWS {
            break;
        }
        for (start, cell) in cells {
            for column in *start..start.saturating_add(cell.repeat) {
                let Ok(col) = u32::try_from(column) else {
                    break;
                };
                if col >= MAX_COLS {
                    break;
                }
                emit_cell(sheet, CellPos { row, col }, cell)?;
            }
        }
    }
    Ok(())
}

fn emit_cell(sheet: &mut SheetBuilder<'_>, pos: CellPos, cell: &CellData) -> Result<(), Error> {
    match (&cell.matrix, &cell.formula) {
        (Matrix::Dropped, _) => {}
        (Matrix::Area { columns, rows }, Some(text)) => {
            // An area that starts inside the sheet and reaches past its edge is cut there.
            let last = CellPos {
                row: pos.row.saturating_add(rows - 1).min(MAX_ROWS - 1),
                col: pos.col.saturating_add(columns - 1).min(MAX_COLS - 1),
            };
            sheet.fill_array(pos, last, text)?;
        }
        _ => {
            let value = match &cell.scalar {
                Scalar::Empty => ValueRef::Empty,
                Scalar::Number(number) => ValueRef::Number(*number),
                Scalar::Bool(flag) => ValueRef::Bool(*flag),
                Scalar::Text(text) => ValueRef::Text(text),
            };
            let formula = cell.formula.as_deref().map(FormulaRef::Text);
            sheet.load_view(pos, &value, formula, cell.stale, None)?;
        }
    }
    Ok(())
}

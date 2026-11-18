//! Streaming parser of a worksheet part.

use std::io::BufRead;
use std::sync::Arc;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use sheetrs_core::{
    CellPos, CellRef, DateSystem, ErrorKind, FormulaRef, Frame, MAX_COLS, MAX_ROWS, SharedMember,
    SheetBuilder, ValueRef,
};

use crate::Error;
use crate::capture::{Tee, without_start_tag};
use crate::lexical;
use crate::xml::{attribute, element_text, raw_attributes, skip_element, unescape_xstring};

/// A `<c>` element while its children are read.
struct PendingCell {
    /// Where the cell goes; `None` when its position is outside the sheet.
    pos: Option<CellPos>,
    kind: Option<String>,
    value: Option<String>,
    inline_text: Option<String>,
    formula: Option<PendingFormula>,
    /// The `s` attribute: the style of the cell.
    style: Option<u32>,
}

struct PendingFormula {
    text: String,
    kind: FormulaKind,
    /// The `ref` attribute, which marks the anchor of a shared group.
    reference: Option<String>,
}

/// What the `t` attribute of a formula makes of it.
#[derive(Clone, Copy, PartialEq, Eq)]
enum FormulaKind {
    /// An ordinary formula, written `normal` or without a type.
    Plain,
    /// A member of a shared group.
    Shared(u32),
    Array,
    /// A data table, which holds no cell of its own.
    DataTable,
    /// A type `LibreOffice` does not know, or a negative group: the formula is dropped and the
    /// cell is a constant.
    Ignored,
}

/// Where the next cell without an address goes.
struct Cursor {
    /// The row of the unaddressed cells; outside the sheet when the row attribute says so.
    row: i64,
    next_row: i64,
    next_col: u32,
}

/// Reads the cells of a worksheet part into `sheet`, and returns the part around its cells; the
/// start of the part is kept only when `from_start` says the part begins in `source`.
pub(crate) fn parse_worksheet<R: BufRead>(
    source: R,
    sheet: &mut SheetBuilder<'_>,
    context: &ValueContext<'_>,
    part: &str,
    from_start: bool,
) -> Result<Frame, Error> {
    let mut reader = Reader::from_reader(Tee::new(source, from_start));
    let mut head = Vec::new();
    // A block of rows is a fragment: its closing tags may have no opening tag in it.
    reader.config_mut().check_end_names = false;
    reader.config_mut().allow_unmatched_ends = true;
    let mut buf = Vec::new();
    let mut scratch = Vec::new();
    let mut cursor = Cursor {
        row: 0,
        next_row: 0,
        next_col: 0,
    };
    let mut pending: Option<PendingCell> = None;
    loop {
        buf.clear();
        match reader
            .read_event_into(&mut buf)
            .map_err(|source| Error::xml(part, source))?
        {
            Event::Start(element) => match (element.local_name().as_ref(), pending.as_mut()) {
                ("sheetData", _) => {
                    head = without_start_tag(reader.get_mut().pause(), element.name().as_ref());
                }
                ("row", _) => start_row(&element, &mut cursor, sheet, part)?,
                ("c", _) => pending = Some(start_cell(&element, &mut cursor, part)?),
                ("f", Some(cell)) => {
                    let (kind, reference) = formula_header(&element, part)?;
                    let text = element_text(&mut reader, &mut scratch, part)?;
                    cell.formula = Some(PendingFormula {
                        text,
                        kind,
                        reference,
                    });
                }
                ("v", Some(cell)) => {
                    cell.value = Some(element_text(&mut reader, &mut scratch, part)?);
                }
                ("is", Some(cell)) => {
                    cell.inline_text = Some(inline_string(&mut reader, part)?);
                }
                _ => {}
            },
            Event::Empty(element) => match (element.local_name().as_ref(), pending.as_mut()) {
                ("sheetData", _) => {
                    head = without_start_tag(reader.get_mut().pause(), element.name().as_ref());
                    reader.get_mut().resume();
                }
                ("row", _) => start_row(&element, &mut cursor, sheet, part)?,
                // A cell without children holds nothing, but it still advances the cursor.
                ("c", _) => {
                    let cell = start_cell(&element, &mut cursor, part)?;
                    record_style(&cell, sheet);
                }
                ("f", Some(cell)) => {
                    let (kind, reference) = formula_header(&element, part)?;
                    cell.formula = Some(PendingFormula {
                        text: String::new(),
                        kind,
                        reference,
                    });
                }
                _ => {}
            },
            Event::End(element) if element.local_name().as_ref() == "c" => {
                if let Some(cell) = pending.take() {
                    record_style(&cell, sheet);
                    finish_cell(&cell, sheet, context)?;
                }
            }
            Event::End(element) if element.local_name().as_ref() == "sheetData" => {
                reader.get_mut().resume();
            }
            // A worksheet cut short keeps the cells read before the cut, as `LibreOffice` does.
            Event::Eof => {
                let tail = reader.get_mut().take();
                return Ok(Frame { head, tail });
            }
            _ => {}
        }
    }
}

/// A row attribute that is not a valid row makes the row `LibreOffice` reads as one outside
/// the sheet: its unaddressed cells are dropped and the next row follows it. The value -1 is
/// the exception: it stands for a row without the attribute.
fn start_row(
    element: &BytesStart<'_>,
    cursor: &mut Cursor,
    sheet: &mut SheetBuilder<'_>,
    part: &str,
) -> Result<(), Error> {
    // The value -1 stands for "no row", as `LibreOffice` reads it.
    cursor.row =
        match attribute(element, "r", part)?.map(|number| sheetrs_io::integer_prefix(&number)) {
            Some(number) if number != -1 => number - 1,
            _ => cursor.next_row,
        };
    cursor.next_row = cursor.row.saturating_add(1);
    cursor.next_col = 0;
    if let Some(row) = u32::try_from(cursor.row).ok().filter(|row| *row < MAX_ROWS) {
        let attributes = raw_attributes(element, &["r", "spans"], part)?;
        if !attributes.is_empty() {
            sheet.set_row_attributes(row, &attributes);
        }
    }
    Ok(())
}

/// Records the style of a cell that has one.
fn record_style(cell: &PendingCell, sheet: &mut SheetBuilder<'_>) {
    if let (Some(pos), Some(style)) = (cell.pos, cell.style.filter(|style| *style != 0)) {
        sheet.set_cell_style(pos, style);
    }
}

/// A cell whose address is not a cell of the sheet is placed like a cell without an address.
fn start_cell(
    element: &BytesStart<'_>,
    cursor: &mut Cursor,
    part: &str,
) -> Result<PendingCell, Error> {
    let addressed = attribute(element, "r", part)?
        .and_then(|address| CellRef::from_a1(&address).ok())
        .map(|cell| CellPos {
            row: cell.row.index,
            col: cell.col.index,
        });
    let pos = addressed.or_else(|| {
        let row = u32::try_from(cursor.row)
            .ok()
            .filter(|row| *row < MAX_ROWS)?;
        (cursor.next_col < MAX_COLS).then_some(CellPos {
            row,
            col: cursor.next_col,
        })
    });
    cursor.next_col = match addressed {
        Some(pos) => pos.col + 1,
        None => cursor.next_col.saturating_add(1),
    };
    Ok(PendingCell {
        pos,
        kind: attribute(element, "t", part)?,
        style: attribute(element, "s", part)?.and_then(|style| style.trim().parse().ok()),
        value: None,
        inline_text: None,
        formula: None,
    })
}

/// The kind of a formula and its `ref` attribute.
fn formula_header(
    element: &BytesStart<'_>,
    part: &str,
) -> Result<(FormulaKind, Option<String>), Error> {
    let kind = formula_kind(element, part)?;
    // Only an anchor is told apart by its `ref`.
    let reference = match kind {
        FormulaKind::Shared(_) | FormulaKind::Array => attribute(element, "ref", part)?,
        _ => None,
    };
    Ok((kind, reference))
}

/// What the `t` and `si` attributes of a formula make of it.
fn formula_kind(element: &BytesStart<'_>, part: &str) -> Result<FormulaKind, Error> {
    Ok(match attribute(element, "t", part)?.as_deref() {
        None | Some("normal") => FormulaKind::Plain,
        Some("shared") => match attribute(element, "si", part)? {
            None => FormulaKind::Plain,
            Some(group) => match u32::try_from(sheetrs_io::integer_prefix(&group)) {
                Ok(group) => FormulaKind::Shared(group),
                Err(_) => FormulaKind::Ignored,
            },
        },
        Some("array") => FormulaKind::Array,
        Some("dataTable") => FormulaKind::DataTable,
        Some(_) => FormulaKind::Ignored,
    })
}

/// The text of an `<is>` element, with rich-text runs joined and phonetic runs dropped.
fn inline_string<R: BufRead>(reader: &mut Reader<R>, part: &str) -> Result<String, Error> {
    let mut text = String::new();
    let mut buf = Vec::new();
    let mut scratch = Vec::new();
    loop {
        buf.clear();
        match reader
            .read_event_into(&mut buf)
            .map_err(|source| Error::xml(part, source))?
        {
            Event::Start(element) => match element.local_name().as_ref() {
                "t" => text.push_str(&unescape_xstring(&element_text(
                    reader,
                    &mut scratch,
                    part,
                )?)),
                "rPh" => skip_element(reader, &element, part)?,
                _ => {}
            },
            Event::End(element) if element.local_name().as_ref() == "is" => {
                return Ok(lexical::normalize_line_breaks(&text).into_owned());
            }
            Event::Eof => return Err(Error::malformed(part, "unexpected end of file")),
            _ => {}
        }
    }
}

fn finish_cell(
    cell: &PendingCell,
    sheet: &mut SheetBuilder<'_>,
    context: &ValueContext<'_>,
) -> Result<(), Error> {
    let Some(pos) = cell.pos else {
        return Ok(());
    };
    let kind = cell.formula.as_ref().map(|formula| formula.kind);
    // A data table holds no cell of its own.
    if kind == Some(FormulaKind::DataTable) {
        return Ok(());
    }
    if let Some(PendingFormula {
        kind: FormulaKind::Array,
        text,
        reference,
    }) = &cell.formula
    {
        return load_array(pos, text, reference.as_deref(), sheet);
    }
    // The text or the `ref` makes the cell an anchor; the group is only kept when an anchor's
    // `ref` covers the anchor itself.
    let member = match &cell.formula {
        Some(PendingFormula {
            kind: FormulaKind::Shared(group),
            text,
            reference,
        }) => Some(SharedMember {
            group: *group,
            anchor_covers_itself: (!text.is_empty() || reference.is_some()).then(|| {
                reference
                    .as_deref()
                    .and_then(lexical::area)
                    .is_some_and(|area| area.contains(pos.row, pos.col))
            }),
        }),
        _ => None,
    };
    let formula = match &cell.formula {
        Some(PendingFormula {
            kind: FormulaKind::Ignored,
            ..
        })
        | None => None,
        Some(PendingFormula { text, kind, .. }) if !text.is_empty() => {
            if let FormulaKind::Shared(group) = kind {
                sheet.set_shared_anchor(*group, pos);
            }
            Some(FormulaRef::Text(text))
        }
        Some(PendingFormula {
            kind: FormulaKind::Shared(group),
            ..
        }) => Some(FormulaRef::Shared { index: *group }),
        Some(_) => None,
    };
    let value = cell_value(
        cell.kind.as_deref(),
        cell.value.as_deref(),
        cell.inline_text.as_deref(),
        formula.is_some(),
        context,
    );
    let implied = if formula.is_none() {
        implied_formula(cell.kind.as_deref(), cell.value.as_deref(), &value)
    } else {
        None
    };
    // Only the non-zero number cached by a plain formula is trusted; everything else is recalculated.
    let trusted = kind == Some(FormulaKind::Plain)
        && matches!(cell.kind.as_deref(), None | Some("n"))
        && !cell
            .formula
            .as_ref()
            .is_some_and(|formula| lexical::recalculated_on_load(&formula.text))
        && matches!(&value, ValueRef::Number(number) if number.is_nan() || number.abs() > 0.0);
    // An implied formula is recalculated on load like any formula that is not a plain
    // non-zero number.
    let stale = (formula.is_some() && !trusted) || implied.is_some();
    let formula = formula.or(implied.map(FormulaRef::Text));
    if value != ValueRef::Empty || formula.is_some() {
        sheet.load_view(pos, &value, formula, stale, member)?;
    }
    Ok(())
}

/// An array formula fills the area of its `ref`, which has to start at the cell that holds
/// the formula; with any other `ref` the cell is dropped, and a range whose end is open fills
/// only its first row.
fn load_array(
    pos: CellPos,
    text: &str,
    reference: Option<&str>,
    sheet: &mut SheetBuilder<'_>,
) -> Result<(), Error> {
    let Some(area) = reference
        .and_then(lexical::area)
        .filter(|area| area.starts_at(pos.row, pos.col))
    else {
        return Ok(());
    };
    let (row, col) = area.end_within(MAX_ROWS, MAX_COLS);
    let row = if area.open_ended { pos.row } else { row };
    sheet.fill_array(pos, CellPos { row, col }, text)?;
    Ok(())
}

/// The formula `LibreOffice` makes of a boolean or error cell that has none: the cell is
/// loaded as `=TRUE()`, `=FALSE()` or the formula that its error text is.
fn implied_formula(
    kind: Option<&str>,
    stored: Option<&str>,
    value: &ValueRef<'_>,
) -> Option<&'static str> {
    match (kind, value) {
        (Some("b"), ValueRef::Bool(true)) => Some("TRUE()"),
        (Some("b"), ValueRef::Bool(false)) => Some("FALSE()"),
        (Some("e"), ValueRef::Error(_)) => lexical::constant_error(stored?).1,
        _ => None,
    }
}

/// What reading a cell value needs to know about the workbook.
pub(crate) struct ValueContext<'a> {
    pub shared_strings: &'a [Arc<str>],
    pub date_system: DateSystem,
}

/// The value of a cell as `LibreOffice` loads it. An unknown cell type holds nothing, and
/// text that does not fit the announced type is read as `LibreOffice` reads it.
fn cell_value<'a>(
    kind: Option<&str>,
    stored: Option<&'a str>,
    inline_text: Option<&'a str>,
    has_formula: bool,
    context: &'a ValueContext<'a>,
) -> ValueRef<'a> {
    // Inline text counts only for inline string cells, and an empty one holds nothing.
    if kind == Some("inlineStr") {
        if stored.is_some() {
            return ValueRef::Empty;
        }
        return inline_text
            .filter(|text| !text.is_empty())
            .map_or(ValueRef::Empty, ValueRef::Text);
    }
    let Some(stored) = stored.filter(|stored| !stored.is_empty()) else {
        return ValueRef::Empty;
    };
    match kind {
        None | Some("n") => ValueRef::Number(lexical::number(stored)),
        Some("s") => lexical::shared_string_index(stored)
            .and_then(|index| context.shared_strings.get(index))
            .filter(|string| !string.is_empty())
            .map_or(ValueRef::Empty, |string| ValueRef::Text(string)),
        Some("str") => ValueRef::Text(stored),
        Some("b") => ValueRef::Bool(lexical::boolean(stored)),
        // The cached error of a formula is kept as stored; it is recalculated on load.
        Some("e") if has_formula => ValueRef::Error(ErrorKind::from_text(stored)),
        Some("e") => ValueRef::Error(lexical::constant_error(stored).0),
        Some("d") => sheetrs_io::iso_date(stored, context.date_system, sheetrs_io::DateRules::XLSX)
            .map_or_else(|| ValueRef::Text(stored), ValueRef::Number),
        Some(_) => ValueRef::Empty,
    }
}

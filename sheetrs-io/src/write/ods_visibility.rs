//! Hiding and showing rows, columns and sheets in an ODS: `table:visibility` of rows and columns,
//! and the `table:display` of the table style of a sheet. The rest of the document is copied as it
//! is.

use super::error::{Context, Error, Result};
use super::part_options;
use super::visibility::{Intervals, SheetVisibility, Visibilities};
use super::{MAX_COL, MAX_ROW, attribute, read_element_body};
use crate::reader::parser_utils::attr_value;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::{HashMap, HashSet};
use std::io::{BufRead, BufReader, BufWriter, Read, Seek, Write};
use zip::{ZipArchive, ZipWriter};

const CONTENT_PART: &str = "content.xml";
const VISIBILITY: &str = "table:visibility";
const ROWS_REPEATED: &str = "table:number-rows-repeated";
const COLUMNS_REPEATED: &str = "table:number-columns-repeated";
const TABLE_STYLE: &str = "table:style-name";

/// A table style of the document and whether it shows its sheet.
struct TableStyle {
    events: Vec<Event<'static>>,
    shown: bool,
}

/// What the document says about its sheets, read before the pass that changes them.
#[derive(Default)]
struct Sheets {
    /// Name and style of each sheet, in order.
    tables: Vec<(String, Option<String>)>,
    styles: HashMap<String, TableStyle>,
}

impl Sheets {
    fn read<R: BufRead>(input: R) -> Result<Self> {
        let mut reader = Reader::from_reader(input);
        let mut found = Self::default();
        let mut depth = 0u32;
        let mut buf = Vec::new();
        loop {
            buf.clear();
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) if e.name().as_ref() == "style:style" => {
                    let start = e.into_owned();
                    let end = start.to_end().into_owned();
                    if attribute(&start, "style:family")?.as_deref() == Some("table") {
                        let name = attribute(&start, "style:name")?.unwrap_or_default();
                        let body = read_element_body(&mut reader, "style:style")?;
                        let shown = !body.iter().any(|event| match event {
                            Event::Start(p) | Event::Empty(p)
                                if p.name().as_ref() == "style:table-properties" =>
                            {
                                attribute(p, "table:display")
                                    .ok()
                                    .flatten()
                                    .is_some_and(|v| v == "false")
                            }
                            _ => false,
                        });
                        let mut events = vec![Event::Start(start)];
                        events.extend(body);
                        events.push(Event::End(end));
                        found.styles.insert(name, TableStyle { events, shown });
                    } else {
                        reader.read_to_end_into(end.name(), &mut Vec::new())?;
                    }
                }
                Event::Start(e) if e.name().as_ref() == "table:table" => {
                    // Only the sheets: a cell may hold a table of its own.
                    if depth == 0 {
                        found.tables.push((
                            attribute(&e, "table:name")?.unwrap_or_default(),
                            attribute(&e, TABLE_STYLE)?,
                        ));
                    }
                    depth += 1;
                }
                Event::Empty(e) if e.name().as_ref() == "table:table" && depth == 0 => {
                    found.tables.push((
                        attribute(&e, "table:name")?.unwrap_or_default(),
                        attribute(&e, TABLE_STYLE)?,
                    ));
                }
                Event::End(e) if e.name().as_ref() == "table:table" => {
                    depth = depth.saturating_sub(1);
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(found)
    }

    fn shown(&self, style: Option<&str>) -> bool {
        style
            .and_then(|name| self.styles.get(name))
            .is_none_or(|style| style.shown)
    }
}

pub(super) fn apply<R: Read + Seek, W: Write + Seek>(
    mut archive: ZipArchive<R>,
    zip_writer: &mut ZipWriter<W>,
    visibilities: &Visibilities,
) -> Result<()> {
    let sheets = if visibilities.states.is_empty() {
        Sheets::default()
    } else {
        Sheets::read(BufReader::new(archive.by_name(CONTENT_PART)?))?
    };
    let mut new_styles = String::new();
    let mut table_styles: HashMap<String, String> = HashMap::new();
    if !visibilities.states.is_empty() {
        let names: Vec<String> = sheets.tables.iter().map(|(name, _)| name.clone()).collect();
        if let Some(unknown) = visibilities.unknown(&names) {
            return Err(Error::SheetNotFound {
                sheet: unknown.to_string(),
                part: CONTENT_PART.to_string(),
            });
        }
        let mut shown_after = 0usize;
        let mut made: HashMap<(Option<String>, bool), String> = HashMap::new();
        for (name, style) in &sheets.tables {
            let shown_now = sheets.shown(style.as_deref());
            let shown_then = visibilities
                .states
                .get(name)
                .map_or(shown_now, |hidden| !hidden);
            shown_after += usize::from(shown_then);
            if shown_then == shown_now {
                continue;
            }
            let key = (style.clone(), shown_then);
            let new_name = made.entry(key).or_insert_with(|| {
                // A name the document does not use, nor an earlier variant of this pass.
                let mut number = 1;
                let new_name = loop {
                    let candidate = format!("sheetrs-sheet-{number}");
                    if !sheets.styles.contains_key(&candidate)
                        && !new_styles.contains(&format!("style:name=\"{candidate}\""))
                    {
                        break candidate;
                    }
                    number += 1;
                };
                new_styles.push_str(&variant_of(
                    style.as_deref().and_then(|s| sheets.styles.get(s)),
                    &new_name,
                    shown_then,
                ));
                new_name
            });
            table_styles.insert(name.clone(), new_name.clone());
        }
        if shown_after == 0 {
            return Err(Error::NoVisibleSheet);
        }
    }
    for i in 0..archive.len() {
        if archive.by_index_raw(i)?.name()? != CONTENT_PART {
            // Raw copies keep `mimetype` first and stored uncompressed, as ODF requires.
            zip_writer.raw_copy_file(archive.by_index_raw(i)?)?;
            continue;
        }
        zip_writer.start_file(CONTENT_PART, part_options())?;
        let seen = rewrite_content(
            BufReader::new(archive.by_index(i)?),
            &mut *zip_writer,
            visibilities,
            &table_styles,
            &new_styles,
        )
        .with_context(|| format!("Failed to modify {CONTENT_PART}"))?;
        if let Some(unknown) = visibilities
            .sheets
            .keys()
            .find(|name| !seen.contains(name.as_str()))
        {
            return Err(Error::SheetNotFound {
                sheet: unknown.clone(),
                part: CONTENT_PART.to_string(),
            });
        }
    }
    Ok(())
}

/// The XML of the table style `name`, as `base` but showing the sheet or not.
fn variant_of(base: Option<&TableStyle>, name: &str, shown: bool) -> String {
    let display = if shown { "true" } else { "false" };
    let mut writer = Writer::new(std::io::Cursor::new(Vec::new()));
    let properties = |from: Option<&BytesStart>| {
        let mut element = BytesStart::new("style:table-properties");
        if let Some(from) = from {
            for attr in from.attributes().flatten() {
                if attr.key.as_ref() != "table:display" {
                    element.push_attribute(attr);
                }
            }
        }
        element.push_attribute(("table:display", display));
        element
    };
    let mut written = |event: Event| writer.write_event(event).expect("a vector takes it");
    match base {
        None => {
            let mut style = BytesStart::new("style:style");
            style.push_attribute(("style:name", name));
            style.push_attribute(("style:family", "table"));
            written(Event::Start(style));
            written(Event::Empty(properties(None)));
            written(Event::End(BytesEnd::new("style:style")));
        }
        Some(base) => {
            let mut has_properties = false;
            for event in &base.events {
                match event {
                    Event::Start(e) if e.name().as_ref() == "style:style" => {
                        let mut style = BytesStart::new("style:style");
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() != "style:name" {
                                style.push_attribute(attr);
                            }
                        }
                        style.push_attribute(("style:name", name));
                        written(Event::Start(style));
                    }
                    Event::Start(e) if e.name().as_ref() == "style:table-properties" => {
                        has_properties = true;
                        written(Event::Start(properties(Some(e))));
                    }
                    Event::Empty(e) if e.name().as_ref() == "style:table-properties" => {
                        has_properties = true;
                        written(Event::Empty(properties(Some(e))));
                    }
                    Event::End(e) if e.name().as_ref() == "style:style" => {
                        if !has_properties {
                            written(Event::Empty(properties(None)));
                        }
                        written(Event::End(e.clone()));
                    }
                    other => written(other.clone()),
                }
            }
        }
    }
    String::from_utf8(writer.into_inner().into_inner()).expect("what was written is text")
}

fn count(element: &BytesStart, key: &str) -> Result<u32> {
    Ok(element
        .try_get_attribute(key)?
        .map(|attr| attr_value(&attr))
        .transpose()?
        .and_then(|n| n.parse::<u32>().ok())
        .unwrap_or(1)
        .max(1))
}

/// `element` with `key` set to `value` (removed when `None`), and with the count of repeats given.
fn restyled(
    element: &BytesStart,
    hidden: Option<bool>,
    repeat: Option<(&str, u32)>,
) -> Result<BytesStart<'static>> {
    let mut out = BytesStart::new(element.name().as_ref().to_string());
    let mut visibility = None;
    for attr in element.attributes() {
        let attr = attr?;
        let key = attr.key.as_ref();
        if key == VISIBILITY {
            visibility = Some(attr_value(&attr)?.into_owned());
        } else if repeat.is_some_and(|(name, _)| key == name) {
            // Written below, in its place.
        } else {
            out.push_attribute((key, &*attr_value(&attr)?));
        }
    }
    let visibility = match (hidden, visibility) {
        // Hidden by a filter stays so: it is hidden.
        (Some(true), Some(v)) if v == "collapse" || v == "filter" => Some(v),
        (Some(true), _) => Some("collapse".to_string()),
        (Some(false), _) => None,
        (None, v) => v,
    };
    if let Some(visibility) = visibility {
        out.push_attribute((VISIBILITY, visibility.as_str()));
    }
    if let Some((name, n)) = repeat.filter(|&(_, n)| n > 1) {
        out.push_attribute((name, n.to_string().as_str()));
    }
    Ok(out)
}

/// The position of the reading inside one sheet.
#[derive(Default)]
struct TableCursor {
    row: u32,
    column: u32,
    /// Whether the columns that no `table:table-column` declares have been written.
    columns_closed: bool,
}

/// An element that is written as a start (the content follows) or on its own.
fn write_as<W: Write>(
    writer: &mut Writer<W>,
    element: BytesStart<'static>,
    empty: bool,
) -> Result<()> {
    writer.write_event(if empty {
        Event::Empty(element)
    } else {
        Event::Start(element)
    })?;
    Ok(())
}

/// A row or a column that is not in the document: `count` of them, with the visibility asked.
fn filler<W: Write>(writer: &mut Writer<W>, row: bool, count: u32, hidden: bool) -> Result<()> {
    let (name, repeated) = if row {
        ("table:table-row", ROWS_REPEATED)
    } else {
        ("table:table-column", COLUMNS_REPEATED)
    };
    let mut element = BytesStart::new(name);
    if hidden {
        element.push_attribute((VISIBILITY, "collapse"));
    }
    if count > 1 {
        element.push_attribute((repeated, count.to_string().as_str()));
    }
    if row {
        writer.write_event(Event::Start(element))?;
        writer.write_event(Event::Empty(BytesStart::new("table:table-cell")))?;
        writer.write_event(Event::End(BytesEnd::new(name)))?;
    } else {
        writer.write_event(Event::Empty(element))?;
    }
    Ok(())
}

/// The hidden runs from `from` on, written after the last position the document declares, with the
/// shown gap between them as a plain run.
fn write_trailing<W: Write>(
    writer: &mut Writer<W>,
    row: bool,
    intervals: &Intervals,
    from: u32,
) -> Result<()> {
    let limit = if row { MAX_ROW - 1 } else { MAX_COL };
    let mut at = from;
    for (first, last) in intervals.hidden_runs(from, limit) {
        if first > at {
            filler(writer, row, first - at, false)?;
        }
        filler(writer, row, last - first + 1, true)?;
        at = last + 1;
    }
    Ok(())
}

/// A row or a column element, cut where the visibility changes inside its repeats.
fn write_repeated<R: BufRead, W: Write>(
    reader: &mut Reader<R>,
    writer: &mut Writer<W>,
    e: &BytesStart<'static>,
    empty: bool,
    intervals: Option<&Intervals>,
    first: u32,
    repeated_key: &str,
) -> Result<()> {
    let repeat = count(e, repeated_key)?;
    let pieces = intervals.map_or_else(
        || vec![(first, first + repeat - 1, None)],
        |intervals| intervals.pieces(first, first + repeat - 1),
    );
    let name = e.name().as_ref().to_string();
    let body = if empty || pieces.len() == 1 {
        None
    } else {
        Some(read_element_body(reader, &name)?)
    };
    if pieces.len() == 1 {
        let (_, _, state) = pieces[0];
        // A single piece keeps the element and its content as they are, read on as the reader goes.
        return write_as(writer, restyled(e, state, None)?, empty);
    }
    for (start, end, state) in pieces {
        let element = restyled(e, state, Some((repeated_key, end - start + 1)))?;
        match &body {
            None => writer.write_event(Event::Empty(element))?,
            Some(body) => {
                writer.write_event(Event::Start(element))?;
                for event in body {
                    writer.write_event(event.borrow())?;
                }
                writer.write_event(Event::End(BytesEnd::new(name.clone())))?;
            }
        }
    }
    Ok(())
}

/// What the pass keeps while it reads `content.xml`.
struct Rewriter<'a> {
    visibilities: &'a Visibilities,
    table_styles: &'a HashMap<String, String>,
    new_styles: &'a str,
    seen: HashSet<String>,
    current: Option<&'a SheetVisibility>,
    cursor: TableCursor,
    /// How many `table:table` are open: 1 is the sheet, more are tables inside cells.
    depth: u32,
}

impl<'a> Rewriter<'a> {
    fn rows(&self) -> Option<&'a Intervals> {
        self.current.map(|s| &s.rows).filter(|i| !i.is_empty())
    }

    fn columns(&self) -> Option<&'a Intervals> {
        self.current.map(|s| &s.columns).filter(|i| !i.is_empty())
    }

    /// The columns no `table:table-column` declares, once the declared ones are over.
    fn close_columns<W: Write>(&mut self, writer: &mut Writer<W>) -> Result<()> {
        if !self.cursor.columns_closed {
            self.cursor.columns_closed = true;
            if let Some(columns) = self.columns() {
                write_trailing(writer, false, columns, self.cursor.column)?;
            }
        }
        Ok(())
    }

    /// What a sheet still owes when it ends: the columns and the rows that it never declared.
    fn finish_table<W: Write>(&mut self, writer: &mut Writer<W>) -> Result<()> {
        self.close_columns(writer)?;
        if let Some(rows) = self.rows() {
            write_trailing(writer, true, rows, self.cursor.row)?;
        }
        Ok(())
    }

    fn start_sheet(&mut self, e: &BytesStart) -> Result<BytesStart<'static>> {
        let name = attribute(e, "table:name")?.unwrap_or_default();
        self.current = self.visibilities.sheets.get(&name);
        self.cursor = TableCursor::default();
        let mut out = e.to_owned();
        if let Some(style) = self.table_styles.get(&name) {
            let mut rebuilt = BytesStart::new(e.name().as_ref().to_string());
            for attr in e.attributes() {
                let attr = attr?;
                if attr.key.as_ref() != TABLE_STYLE {
                    rebuilt.push_attribute(attr);
                }
            }
            rebuilt.push_attribute((TABLE_STYLE, style.as_str()));
            out = rebuilt;
        }
        self.seen.insert(name);
        Ok(out)
    }

    fn element<R: BufRead, W: Write>(
        &mut self,
        reader: &mut Reader<R>,
        writer: &mut Writer<W>,
        e: BytesStart<'static>,
        empty: bool,
    ) -> Result<()> {
        match e.name().as_ref() {
            "table:table" => {
                if self.depth == 0 {
                    let sheet = self.start_sheet(&e)?;
                    if empty {
                        // A sheet with nothing in it still gets what it was asked for.
                        writer.write_event(Event::Start(sheet))?;
                        self.finish_table(writer)?;
                        writer.write_event(Event::End(BytesEnd::new("table:table")))?;
                        return Ok(());
                    }
                    write_as(writer, sheet, false)?;
                } else {
                    write_as(writer, e, empty)?;
                }
                if !empty {
                    self.depth += 1;
                }
            }
            "table:table-column" if self.depth == 1 => {
                let repeat = count(&e, COLUMNS_REPEATED)?;
                let columns = self.columns();
                write_repeated(
                    reader,
                    writer,
                    &e,
                    empty,
                    columns,
                    self.cursor.column,
                    COLUMNS_REPEATED,
                )?;
                self.cursor.column += repeat;
            }
            "table:table-row" if self.depth == 1 => {
                self.close_columns(writer)?;
                let repeat = count(&e, ROWS_REPEATED)?;
                let rows = self.rows();
                write_repeated(
                    reader,
                    writer,
                    &e,
                    empty,
                    rows,
                    self.cursor.row,
                    ROWS_REPEATED,
                )?;
                self.cursor.row += repeat;
            }
            "table:table-rows" | "table:table-header-rows" | "table:table-row-group"
                if self.depth == 1 =>
            {
                self.close_columns(writer)?;
                write_as(writer, e, empty)?;
            }
            "office:automatic-styles" if empty && !self.new_styles.is_empty() => {
                writer.write_event(Event::Start(e))?;
                writer.get_mut().write_all(self.new_styles.as_bytes())?;
                writer.write_event(Event::End(BytesEnd::new("office:automatic-styles")))?;
            }
            _ => write_as(writer, e, empty)?,
        }
        Ok(())
    }

    fn end<W: Write>(&mut self, writer: &mut Writer<W>, e: BytesEnd<'static>) -> Result<()> {
        match e.name().as_ref() {
            "table:table" => {
                if self.depth == 1 {
                    self.finish_table(writer)?;
                }
                self.depth = self.depth.saturating_sub(1);
            }
            "office:automatic-styles" => {
                writer.get_mut().write_all(self.new_styles.as_bytes())?;
            }
            _ => {}
        }
        writer.write_event(Event::End(e))?;
        Ok(())
    }
}

/// Rewrite `content.xml`; the names of the sheets it holds.
fn rewrite_content<R: BufRead, W: Write>(
    input: R,
    output: W,
    visibilities: &Visibilities,
    table_styles: &HashMap<String, String>,
    new_styles: &str,
) -> Result<HashSet<String>> {
    let mut reader = Reader::from_reader(input);
    let mut writer = Writer::new(BufWriter::new(output));
    let mut state = Rewriter {
        visibilities,
        table_styles,
        new_styles,
        seen: HashSet::new(),
        current: None,
        cursor: TableCursor::default(),
        depth: 0,
    };
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => state.element(&mut reader, &mut writer, e.into_owned(), false)?,
            Event::Empty(e) => state.element(&mut reader, &mut writer, e.into_owned(), true)?,
            Event::End(e) => state.end(&mut writer, e.into_owned())?,
            Event::Eof => break,
            event => writer.write_event(event)?,
        }
    }
    writer.into_inner().flush()?;
    Ok(state.seen)
}

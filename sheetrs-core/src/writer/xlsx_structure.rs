//! The parts of an XLSX worksheet that a deletion of rows or columns changes besides its cells:
//! each element that holds a range or a formula, as `xlsx_writer::transform_sheet` meets it.

use super::structure::{Bands, StructureEdits, shape_depends_on_sizes};
use super::{attribute, read_element_body};
use crate::reader::parser_utils::{
    attr_value, column_letters, parse_cell_range, parse_cell_ref, push_general_ref,
};
use anyhow::{Result, bail};
use quick_xml::events::{BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use sheetrs_formula::edit::Edit;
use std::io::{BufRead, Write};

/// What a command does to one worksheet.
pub(super) struct SheetStructure<'a> {
    pub plan: &'a StructureEdits,
    pub edits: &'a [Edit<'a>],
    pub sheet: String,
    pub rows: Bands,
    pub cols: Bands,
    /// The formulas lose their cached results so the application that opens the file computes
    /// them (an ODS has no marker for it, as an XLSX has `fullCalcOnLoad`).
    pub recalculate: bool,
}

impl SheetStructure<'_> {
    /// Whether the sheet has rows or columns deleted.
    pub fn changes_layout(&self) -> bool {
        !self.rows.is_empty() || !self.cols.is_empty()
    }

    /// The text of a formula of this sheet after the plan, `None` when it does not change.
    pub fn rewrite(&self, formula: &str) -> Result<Option<String>> {
        if formula.is_empty()
            || (!self.changes_layout() && !self.plan.may_reach_edited_sheet(formula))
        {
            return Ok(None);
        }
        self.plan.rewrite_formula(self.edits, formula, &self.sheet)
    }

    /// `A1`, `A1:B2`, `C:D` or `3:5` once the plan is applied, `None` when no cell of it remains.
    /// A text that is none of them (names) is returned as it is.
    pub fn range_text(&self, range: &str) -> Option<String> {
        use sheetrs_formula::addr::{col_name, parse_col, parse_row};
        if let Some((a, b)) = range.split_once(':') {
            if let (Some((c0, _)), Some((c1, _))) = (parse_col(a), parse_col(b)) {
                let (c0, c1) = self.cols.span(c0.min(c1), c0.max(c1))?;
                return Some(format!("{}:{}", col_name(c0), col_name(c1)));
            }
            if let (Some((r0, _)), Some((r1, _))) = (parse_row(a), parse_row(b)) {
                let (r0, r1) = self.rows.span(r0.min(r1), r0.max(r1))?;
                return Some(format!("{}:{}", r0 + 1, r1 + 1));
            }
        }
        let Some((r0, c0, r1, c1)) = parse_cell_range(range).or_else(|| {
            parse_cell_ref(range)
                .filter(|_| range.bytes().all(|b| b.is_ascii_alphanumeric()))
                .map(|(r, c)| (r, c, r, c))
        }) else {
            return Some(range.to_string());
        };
        let (r0, c0, r1, c1) = self.plan.range(&self.sheet, r0, c0, r1, c1)?;
        Some(if (r0, c0) == (r1, c1) {
            format!("{}{}", column_letters(c0), r0 + 1)
        } else {
            format!(
                "{}{}:{}{}",
                column_letters(c0),
                r0 + 1,
                column_letters(c1),
                r1 + 1
            )
        })
    }

    /// A list of ranges separated by spaces (`sqref`) once the plan is applied, `None` when no
    /// cell of any remains.
    pub fn sqref_text(&self, sqref: &str) -> Option<String> {
        let kept: Vec<String> = sqref
            .split_whitespace()
            .filter_map(|range| self.range_text(range))
            .collect();
        (!kept.is_empty()).then(|| kept.join(" "))
    }
}

/// A copy of `element` with each attribute mapped by `map`: `None` keeps it, `Some(None)` drops
/// it and `Some(Some(value))` sets it.
pub(super) fn map_attributes(
    element: &BytesStart,
    mut map: impl FnMut(&str, &str) -> Result<Option<Option<String>>>,
) -> Result<BytesStart<'static>> {
    let mut rebuilt = BytesStart::new(element.name().as_ref().to_string());
    for attr in element.attributes() {
        let attr = attr?;
        let key = attr.key.as_ref().to_string();
        let value = attr_value(&attr)?.into_owned();
        match map(&key, &value)? {
            None => rebuilt.push_attribute((key.as_str(), value.as_str())),
            Some(None) => {}
            Some(Some(new)) => rebuilt.push_attribute((key.as_str(), new.as_str())),
        }
    }
    Ok(rebuilt)
}

/// The text content of events, with its entity references resolved.
pub(super) fn text_of(body: &[Event]) -> Result<String> {
    let mut text = String::new();
    for event in body {
        match event {
            Event::Text(t) => text.push_str(t),
            Event::CData(t) => text.push_str(t),
            Event::GeneralRef(r) => push_general_ref(&mut text, r)?,
            _ => {}
        }
    }
    Ok(text)
}

/// The name of an element without its namespace prefix.
fn local_name(e: &BytesStart) -> String {
    let name = e.name();
    let name = name.as_ref();
    name.rsplit(':').next().unwrap_or(name).to_string()
}

/// The conditional formats, validations and sparklines of an extension (`x14:`) with the ranges
/// (`xm:sqref`) and formulas (`xm:f`) they hold adjusted, and those whose cells are all gone dropped.
fn rewrite_extensions(body: &[Event<'static>], s: &SheetStructure) -> Result<Vec<Event<'static>>> {
    let mut out = Vec::with_capacity(body.len());
    let mut i = 0;
    while i < body.len() {
        match &body[i] {
            Event::Start(e)
                if ["dataValidation", "conditionalFormatting", "sparkline"]
                    .contains(&local_name(e).as_str()) =>
            {
                let (mut depth, mut end) = (0u32, i);
                for (j, event) in body.iter().enumerate().skip(i) {
                    match event {
                        Event::Start(_) => depth += 1,
                        Event::End(_) => {
                            depth -= 1;
                            if depth == 0 {
                                end = j;
                                break;
                            }
                        }
                        _ => {}
                    }
                }
                out.extend(rewrite_extension_element(&body[i..=end], s)?.unwrap_or_default());
                i = end + 1;
            }
            // The count would be out of date once elements go.
            Event::Start(e) if local_name(e) == "dataValidations" => {
                out.push(Event::Start(map_attributes(e, |key, _| {
                    Ok((key == "count").then_some(None))
                })?));
                i += 1;
            }
            event => {
                out.push(event.clone());
                i += 1;
            }
        }
    }
    Ok(prune_empty(out))
}

/// The index of the end event that closes the element that starts at `start`.
fn closing_event(events: &[Event<'static>], start: usize) -> usize {
    let mut depth = 0u32;
    for (j, event) in events.iter().enumerate().skip(start) {
        match event {
            Event::Start(_) => depth += 1,
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    return j;
                }
            }
            _ => {}
        }
    }
    events.len() - 1
}

/// The events without the wrappers that the removal of their content left empty: the schema
/// asks for at least one child in each. A group of sparklines goes when it has no `sparklines`.
fn prune_empty(mut events: Vec<Event<'static>>) -> Vec<Event<'static>> {
    loop {
        let before = events.len();
        let mut out = Vec::with_capacity(before);
        let mut i = 0;
        while i < events.len() {
            if let Event::Start(e) = &events[i] {
                let name = local_name(e);
                let wrapper = matches!(
                    name.as_str(),
                    "dataValidations"
                        | "conditionalFormattings"
                        | "sparklines"
                        | "sparklineGroups"
                        | "sparklineGroup"
                        | "ext"
                );
                if wrapper {
                    let end = closing_event(&events, i);
                    let inner = &events[i + 1..end];
                    let has = |wanted: Option<&str>| {
                        inner.iter().any(|event| match event {
                            Event::Start(x) | Event::Empty(x) => {
                                wanted.is_none_or(|wanted| local_name(x) == wanted)
                            }
                            _ => false,
                        })
                    };
                    let empty = if name == "sparklineGroup" {
                        !has(Some("sparklines"))
                    } else {
                        !has(None)
                    };
                    if empty {
                        i = end + 1;
                        continue;
                    }
                }
            }
            out.push(events[i].clone());
            i += 1;
        }
        events = out;
        if events.len() == before {
            return events;
        }
    }
}

/// One element of an extension with its ranges and formulas adjusted, `None` when its cells go.
fn rewrite_extension_element(
    element: &[Event<'static>],
    s: &SheetStructure,
) -> Result<Option<Vec<Event<'static>>>> {
    let mut out = Vec::with_capacity(element.len());
    let mut i = 0;
    while i < element.len() {
        let kind = match &element[i] {
            Event::Start(e) if e.name().as_ref().ends_with(":sqref") => Some(true),
            Event::Start(e) if e.name().as_ref().ends_with(":f") => Some(false),
            _ => None,
        };
        let Some(is_sqref) = kind else {
            out.push(element[i].clone());
            i += 1;
            continue;
        };
        let end = element[i..]
            .iter()
            .position(|event| matches!(event, Event::End(_)))
            .map_or(element.len() - 1, |offset| i + offset);
        let text = text_of(&element[i + 1..end])?;
        let new = if is_sqref {
            match s.sqref_text(&text) {
                Some(new) => new,
                None => return Ok(None),
            }
        } else {
            s.rewrite(&text)?.unwrap_or(text)
        };
        out.push(element[i].clone());
        out.push(Event::Text(BytesText::new(&new).into_owned()));
        out.push(element[end].clone());
        i = end + 1;
    }
    Ok(Some(out))
}

/// The criteria of a filter (`filterColumn`, numbered from the first column of the range, `c0`)
/// once the columns are gone: those of a deleted column go, the others take the number their
/// column has in the range that starts at `new_c0`.
fn renumber_filter_columns(
    body: &[Event<'static>],
    c0: u32,
    new_c0: u32,
    s: &SheetStructure,
) -> Result<Vec<Event<'static>>> {
    let mut out = Vec::with_capacity(body.len());
    let mut i = 0;
    while i < body.len() {
        let (Event::Start(e) | Event::Empty(e)) = &body[i] else {
            out.push(body[i].clone());
            i += 1;
            continue;
        };
        // The sort of a filter keeps ranges of its own.
        if matches!(e.name().as_ref(), "sortState" | "sortCondition") {
            let range = attribute(e, "ref")?.unwrap_or_default();
            match s.range_text(&range) {
                Some(new) => {
                    let mapped =
                        map_attributes(e, |key, _| Ok((key == "ref").then(|| Some(new.clone()))))?;
                    out.push(match &body[i] {
                        Event::Start(_) => Event::Start(mapped),
                        _ => Event::Empty(mapped),
                    });
                    i += 1;
                }
                None if matches!(body[i], Event::Start(_)) => i = closing_event(body, i) + 1,
                None => i += 1,
            }
            continue;
        }
        if e.name().as_ref() != "filterColumn" {
            out.push(body[i].clone());
            i += 1;
            continue;
        }
        let end = if matches!(body[i], Event::Start(_)) {
            closing_event(body, i)
        } else {
            i
        };
        let id = attribute(e, "colId")?.and_then(|id| id.parse::<u32>().ok());
        match id.map(|id| s.cols.moved(c0 + id)) {
            // A column that is gone takes its criteria with it.
            Some(None) => {}
            Some(Some(new_col)) => {
                let renumbered = map_attributes(e, |key, _| {
                    Ok((key == "colId").then(|| Some((new_col - new_c0).to_string())))
                })?;
                out.push(match &body[i] {
                    Event::Start(_) => Event::Start(renumbered),
                    _ => Event::Empty(renumbered),
                });
                out.extend(body[i + 1..=end].iter().cloned());
            }
            None => out.extend(body[i..=end].iter().cloned()),
        }
        i = end + 1;
    }
    Ok(out)
}

/// What the main loop does with an element this module knows.
pub(super) enum Handled {
    /// The element was written (or dropped); read the next event.
    Done,
    /// Nothing to change: the loop writes it as it is.
    No,
}

/// Rewrite the element `start` of a worksheet when it holds ranges or formulas that the plan
/// changes. `is_start` tells an element with content from an empty one: when the element is kept,
/// its content follows through the main loop; when it goes, its content goes with it.
pub(super) fn handle_element<R: BufRead, W: Write>(
    reader: &mut Reader<R>,
    writer: &mut Writer<W>,
    start: &BytesStart,
    is_start: bool,
    s: &SheetStructure,
) -> Result<Handled> {
    let name = start.name().as_ref().to_string();
    let write = |writer: &mut Writer<W>, element: BytesStart<'static>| -> Result<()> {
        writer.write_event(if is_start {
            Event::Start(element)
        } else {
            Event::Empty(element)
        })?;
        Ok(())
    };
    let skip_content = |reader: &mut Reader<R>| -> Result<()> {
        if is_start {
            read_element_body(reader, &name)?;
        }
        Ok(())
    };
    let layout = s.changes_layout();
    match name.as_str() {
        "dimension" if layout => {
            let element =
                map_attributes(start, |key, value| {
                    Ok((key == "ref")
                        .then(|| Some(s.range_text(value).unwrap_or_else(|| "A1".into()))))
                })?;
            write(writer, element)?;
        }
        "col" if layout => {
            // `min` and `max` are 1-based columns.
            let (Some(min), Some(max)) = (attribute(start, "min")?, attribute(start, "max")?)
            else {
                return Ok(Handled::No);
            };
            let (min, max) = (min.parse::<u32>()? - 1, max.parse::<u32>()? - 1);
            match s.cols.span(min, max) {
                Some((first, last)) => {
                    let element = map_attributes(start, |key, _| {
                        Ok(match key {
                            "min" => Some(Some((first + 1).to_string())),
                            "max" => Some(Some((last + 1).to_string())),
                            _ => None,
                        })
                    })?;
                    write(writer, element)?;
                }
                None => skip_content(reader)?,
            }
        }
        "mergeCells" | "dataValidations" if layout => {
            // The count would be out of date once elements go.
            let element = map_attributes(start, |key, _| Ok((key == "count").then_some(None)))?;
            write(writer, element)?;
        }
        "mergeCell" if layout => {
            let range = attribute(start, "ref")?.unwrap_or_default();
            // A merged area of one cell is not an area.
            let kept = s.range_text(&range).filter(|r| r.contains(':'));
            // And an area that stays needs its first cell.
            if let (Some(_), Some((row, col, ..))) = (&kept, parse_cell_range(&range))
                && (s.rows.contains(row) || s.cols.contains(col))
            {
                bail!(
                    "a deletion in sheet '{}' removes the first cell of the merged area {range}, which stays",
                    s.sheet
                );
            }
            match kept {
                Some(new) => {
                    let element = map_attributes(start, |key, _| {
                        Ok((key == "ref").then(|| Some(new.clone())))
                    })?;
                    write(writer, element)?;
                }
                None => skip_content(reader)?,
            }
        }
        "conditionalFormatting" | "dataValidation" | "ignoredError" | "protectedRange"
            if layout =>
        {
            let sqref = attribute(start, "sqref")?.unwrap_or_default();
            match s.sqref_text(&sqref) {
                Some(new) => {
                    let element = map_attributes(start, |key, _| {
                        Ok((key == "sqref").then(|| Some(new.clone())))
                    })?;
                    write(writer, element)?;
                }
                None => skip_content(reader)?,
            }
        }
        "hyperlink" => {
            // A cell that is gone takes its link with it.
            let range = attribute(start, "ref")?.unwrap_or_default();
            let Some(new_ref) = (if layout {
                s.range_text(&range)
            } else {
                Some(range)
            }) else {
                skip_content(reader)?;
                return Ok(Handled::Done);
            };
            let element = map_attributes(start, |key, value| {
                Ok(match key {
                    "ref" => Some(Some(new_ref.clone())),
                    "location" => s.rewrite(value)?.map(Some),
                    _ => None,
                })
            })?;
            write(writer, element)?;
        }
        "autoFilter" | "sortState" if layout => {
            let range = attribute(start, "ref")?.unwrap_or_default();
            let first_col = parse_cell_range(&range).map(|(_, c0, ..)| c0);
            match s.range_text(&range) {
                Some(new) => {
                    let element = map_attributes(start, |key, _| {
                        Ok((key == "ref").then(|| Some(new.clone())))
                    })?;
                    match (name.as_str(), is_start, first_col, parse_cell_range(&new)) {
                        ("autoFilter", true, Some(c0), Some((_, new_c0, ..))) => {
                            let body = read_element_body(reader, &name)?;
                            writer.write_event(Event::Start(element))?;
                            for event in &renumber_filter_columns(&body, c0, new_c0, s)? {
                                writer.write_event(event.borrow())?;
                            }
                            writer.write_event(Event::End(start.to_end()))?;
                        }
                        _ => write(writer, element)?,
                    }
                }
                None => skip_content(reader)?,
            }
        }
        "sortCondition" if layout => {
            let range = attribute(start, "ref")?.unwrap_or_default();
            match s.range_text(&range) {
                Some(new) => {
                    let element = map_attributes(start, |key, _| {
                        Ok((key == "ref").then(|| Some(new.clone())))
                    })?;
                    write(writer, element)?;
                }
                None => skip_content(reader)?,
            }
        }
        "extLst" if is_start => {
            let body = read_element_body(reader, "extLst")?;
            let body = rewrite_extensions(&body, s)?;
            // An extension list with nothing left is not written.
            if body
                .iter()
                .any(|event| matches!(event, Event::Start(_) | Event::Empty(_)))
            {
                writer.write_event(Event::Start(start.borrow()))?;
                for event in &body {
                    writer.write_event(event.borrow())?;
                }
                writer.write_event(Event::End(start.to_end()))?;
            }
        }
        "pane" if layout => {
            let element = map_attributes(start, |key, value| {
                Ok((key == "topLeftCell")
                    .then(|| Some(s.range_text(value).unwrap_or_else(|| "A1".into()))))
            })?;
            write(writer, element)?;
        }
        "selection" if layout => {
            let element = map_attributes(start, |key, value| {
                Ok(match key {
                    "activeCell" => Some(Some(s.range_text(value).unwrap_or_else(|| "A1".into()))),
                    "sqref" => Some(Some(s.sqref_text(value).unwrap_or_else(|| "A1".into()))),
                    _ => None,
                })
            })?;
            write(writer, element)?;
        }
        "legacyDrawingHF" | "tableParts" if layout => {
            bail!(
                "sheet '{}' has tables or header and footer images, whose positions are not adjusted yet",
                s.sheet
            );
        }
        "formula" | "formula1" | "formula2" if is_start => {
            let body = read_element_body(reader, &name)?;
            let text = text_of(&body)?;
            let new = s.rewrite(&text)?.unwrap_or(text);
            writer.write_event(Event::Start(BytesStart::new(name.clone())))?;
            writer.write_event(Event::Text(BytesText::new(&new)))?;
            writer.write_event(Event::End(quick_xml::events::BytesEnd::new(name)))?;
        }
        _ => return Ok(Handled::No),
    }
    Ok(Handled::Done)
}

/// The comments of a sheet (`xl/commentsN.xml`, `xl/threadedComments/…`, the elements named
/// `element`) with the cells they sit on moved; those on deleted cells go.
pub(super) fn rewrite_comments(xml: &str, element: &str, s: &SheetStructure) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(std::io::Cursor::new(Vec::new()));
    let moved = |e: &BytesStart| -> Result<Option<BytesStart<'static>>> {
        let at = attribute(e, "ref")?.unwrap_or_default();
        s.range_text(&at)
            .map(|new| map_attributes(e, |key, _| Ok((key == "ref").then(|| Some(new.clone())))))
            .transpose()
    };
    loop {
        match reader.read_event()? {
            Event::Start(e) if e.name().as_ref() == element => match moved(&e)? {
                Some(moved) => writer.write_event(Event::Start(moved))?,
                None => {
                    let end = e.to_end().into_owned();
                    reader.read_to_end(end.name())?;
                }
            },
            Event::Empty(e) if e.name().as_ref() == element => {
                if let Some(moved) = moved(&e)? {
                    writer.write_event(Event::Empty(moved))?;
                }
            }
            Event::Eof => break,
            event => writer.write_event(event)?,
        }
    }
    Ok(String::from_utf8(writer.into_inner().into_inner())?)
}

/// The shapes of a comments drawing (`xl/drawings/vmlDrawingN.vml`): each note moves with its
/// cell, and the note of a deleted cell goes.
pub(super) fn rewrite_vml(xml: &str, s: &SheetStructure) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(std::io::Cursor::new(Vec::new()));
    loop {
        match reader.read_event()? {
            Event::Start(e) if e.name().as_ref() == "v:shape" => {
                let end = e.to_end().into_owned();
                let body = read_element_body(&mut reader, "v:shape")?;
                match moved_note(&body, s)? {
                    None => {
                        writer.write_event(Event::Start(e))?;
                        for event in &body {
                            writer.write_event(event.borrow())?;
                        }
                        writer.write_event(Event::End(end))?;
                    }
                    Some(None) => {}
                    Some(Some(events)) => {
                        writer.write_event(Event::Start(e))?;
                        for event in &events {
                            writer.write_event(event.borrow())?;
                        }
                        writer.write_event(Event::End(end))?;
                    }
                }
            }
            Event::Eof => break,
            event => writer.write_event(event)?,
        }
    }
    Ok(String::from_utf8(writer.into_inner().into_inner())?)
}

/// The content of a note shape once its cell moves: `None` when the shape is not a note,
/// `Some(None)` when its cell is deleted.
fn moved_note(
    body: &[Event<'static>],
    s: &SheetStructure,
) -> Result<Option<Option<Vec<Event<'static>>>>> {
    let mut is_note = false;
    let (mut row, mut col) = (None, None);
    let mut inside = "";
    for event in body {
        match event {
            Event::Start(e) => {
                inside = match e.name().as_ref() {
                    "x:Row" => "row",
                    "x:Column" => "col",
                    "x:ClientData" => {
                        is_note = attribute(e, "ObjectType")?.as_deref() == Some("Note");
                        ""
                    }
                    _ => "",
                };
            }
            Event::Text(t) => match inside {
                "row" => row = t.trim().parse::<u32>().ok(),
                "col" => col = t.trim().parse::<u32>().ok(),
                _ => {}
            },
            Event::End(_) => inside = "",
            _ => {}
        }
    }
    let (true, Some(row), Some(col)) = (is_note, row, col) else {
        return Ok(None);
    };
    let Some((new_row, new_col, ..)) = s.plan.range(&s.sheet, row, col, row, col) else {
        return Ok(Some(None));
    };
    let (drow, dcol) = (
        i64::from(new_row) - i64::from(row),
        i64::from(new_col) - i64::from(col),
    );
    let mut out = Vec::with_capacity(body.len());
    inside = "";
    for event in body {
        match event {
            Event::Start(e) => {
                inside = match e.name().as_ref() {
                    "x:Row" => "row",
                    "x:Column" => "col",
                    "x:Anchor" => "anchor",
                    _ => "",
                };
                out.push(event.clone());
            }
            Event::Text(_) if inside == "row" => {
                out.push(Event::Text(
                    BytesText::new(&new_row.to_string()).into_owned(),
                ));
            }
            Event::Text(_) if inside == "col" => {
                out.push(Event::Text(
                    BytesText::new(&new_col.to_string()).into_owned(),
                ));
            }
            Event::Text(t) if inside == "anchor" => {
                // Left column, its offset, top row, its offset, and the same for the right end.
                let numbers: Vec<i64> =
                    t.split(',').filter_map(|n| n.trim().parse().ok()).collect();
                if let [a, b, c, d, e, f, g, h] = numbers[..] {
                    let text = format!(
                        "{}, {b}, {}, {d}, {}, {f}, {}, {h}",
                        a + dcol,
                        c + drow,
                        e + dcol,
                        g + drow
                    );
                    out.push(Event::Text(BytesText::new(&text).into_owned()));
                } else {
                    out.push(event.clone());
                }
            }
            Event::End(_) => {
                inside = "";
                out.push(event.clone());
            }
            event => out.push(event.clone()),
        }
    }
    Ok(Some(Some(out)))
}

/// The anchors of the shapes of a drawing (`xl/drawings/drawingN.xml`: images, charts, shapes)
/// once the rows and columns of its sheet are deleted, as LibreOffice moves them. The top left
/// corner goes to the place of its cell, or keeps its index when the cell goes. A shape that
/// resizes with its cells (`editAs="twoCell"`, which is the default) has its bottom right corner
/// follow its cell the same way; any other keeps its size in rows and columns.
pub(super) fn rewrite_drawing(xml: &str, s: &SheetStructure) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(std::io::Cursor::new(Vec::new()));
    loop {
        match reader.read_event()? {
            Event::Start(e)
                if matches!(e.name().as_ref(), "xdr:twoCellAnchor" | "xdr:oneCellAnchor") =>
            {
                let name = e.name().as_ref().to_string();
                let end = e.to_end().into_owned();
                let resizes = name == "xdr:twoCellAnchor"
                    && attribute(&e, "editAs")?.is_none_or(|edit| edit == "twoCell");
                let body = read_element_body(&mut reader, &name)?;
                writer.write_event(Event::Start(e))?;
                for event in anchored(&body, s, resizes)? {
                    writer.write_event(event.borrow())?;
                }
                writer.write_event(Event::End(end))?;
            }
            Event::Eof => break,
            event => writer.write_event(event)?,
        }
    }
    Ok(String::from_utf8(writer.into_inner().into_inner())?)
}

/// The content of an anchor with its `from` and `to` cells moved; `resizes` is whether the shape
/// resizes with its cells.
fn anchored(
    body: &[Event<'static>],
    s: &SheetStructure,
    resizes: bool,
) -> Result<Vec<Event<'static>>> {
    // The cell of each corner, as `(row, col)`: first `from`, then `to`.
    let (mut corner, mut field) = ("", "");
    let mut from = (None, None);
    let mut to = (None, None);
    for event in body {
        match event {
            Event::Start(e) => match e.name().as_ref() {
                "xdr:from" => corner = "from",
                "xdr:to" => corner = "to",
                "xdr:row" => field = "row",
                "xdr:col" => field = "col",
                _ => {}
            },
            Event::Text(t) => {
                let n = t.trim().parse::<u32>().ok();
                match (corner, field) {
                    ("from", "row") => from.0 = n,
                    ("from", "col") => from.1 = n,
                    ("to", "row") => to.0 = n,
                    ("to", "col") => to.1 = n,
                    _ => {}
                }
            }
            Event::End(e) if matches!(e.name().as_ref(), "xdr:row" | "xdr:col") => field = "",
            Event::End(e) if matches!(e.name().as_ref(), "xdr:from" | "xdr:to") => corner = "",
            _ => {}
        }
    }
    let (Some(from_row), Some(from_col)) = from else {
        return Ok(body.to_vec());
    };
    if s.rows.deleted_below_a_deleted(from_row) || s.cols.deleted_below_a_deleted(from_col) {
        return Err(shape_depends_on_sizes(&s.sheet));
    }
    let new_row = s.rows.anchored_at(from_row);
    let new_col = s.cols.anchored_at(from_col);
    let (mut corner, mut field) = ("", "");
    let mut out = Vec::with_capacity(body.len());
    for event in body {
        match event {
            Event::Start(e) => {
                match e.name().as_ref() {
                    "xdr:from" => corner = "from",
                    "xdr:to" => corner = "to",
                    "xdr:row" => field = "row",
                    "xdr:col" => field = "col",
                    _ => {}
                }
                out.push(event.clone());
            }
            Event::Text(t) => {
                let moved = match (corner, field) {
                    ("from", "row") => Some(new_row),
                    ("from", "col") => Some(new_col),
                    ("to", "row") if resizes => to.0.map(|r| s.rows.anchored_at(r)),
                    ("to", "col") if resizes => to.1.map(|c| s.cols.anchored_at(c)),
                    // The size in rows and columns is what stays.
                    ("to", "row") => to.0.map(|r| new_row + r.saturating_sub(from_row)),
                    ("to", "col") => to.1.map(|c| new_col + c.saturating_sub(from_col)),
                    _ => None,
                };
                match moved {
                    Some(n) => out.push(Event::Text(BytesText::new(&n.to_string()).into_owned())),
                    None => out.push(Event::Text(t.clone().into_owned())),
                }
            }
            Event::End(e) => {
                match e.name().as_ref() {
                    "xdr:row" | "xdr:col" => field = "",
                    "xdr:from" | "xdr:to" => corner = "",
                    _ => {}
                }
                out.push(event.clone());
            }
            event => out.push(event.clone()),
        }
    }
    Ok(out)
}

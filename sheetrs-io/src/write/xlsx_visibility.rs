//! Hiding and showing rows, columns and sheets in an XLSX: `hidden` of `<row>` and `<col>`, `state`
//! of `<sheet>`. The rest of the file is copied as it is.

use super::error::{Context, Error, Result};
use super::part_options;
use super::visibility::{Intervals, Visibilities};
use super::xlsx_writer::{
    SheetEntry, WORKBOOK_PART, WORKBOOK_RELS_PART, parse_relationship_targets, parse_sheets,
    read_part,
};
use super::{MAX_COL, MAX_ROW, attribute};
use crate::reader::parser_utils::attr_value;
use crate::reader::xlsx_parser::workbook_part_name;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::io::{BufRead, BufReader, BufWriter, Read, Seek, Write};
use zip::{ZipArchive, ZipWriter};

/// The width a column that was not declared has, in characters (the one of Excel's default font).
const DEFAULT_COLUMN_WIDTH: &str = "9.140625";

/// What the pass does to the worksheet part of one sheet.
#[derive(Default)]
struct SheetPlan<'a> {
    rows: Option<&'a Intervals>,
    columns: Option<&'a Intervals>,
    /// `Some(false)`: the sheet is no longer the selected tab; `Some(true)`: it is now.
    tab_selected: Option<bool>,
}

impl SheetPlan<'_> {
    fn is_empty(&self) -> bool {
        self.rows.is_none_or(Intervals::is_empty)
            && self.columns.is_none_or(Intervals::is_empty)
            && self.tab_selected.is_none()
    }
}

/// The state of a `<sheet>` as the file says it: `None` for a visible one.
fn current_state(entry: &BytesStart) -> Result<Option<String>> {
    Ok(attribute(entry, "state")?.filter(|state| state != "visible"))
}

pub(super) fn apply<R: Read + Seek, W: Write + Seek>(
    mut archive: ZipArchive<R>,
    zip_writer: &mut ZipWriter<W>,
    visibilities: &Visibilities,
) -> Result<()> {
    let workbook_xml = read_part(&mut archive, WORKBOOK_PART)?;
    let rels_xml = read_part(&mut archive, WORKBOOK_RELS_PART)?;
    let sheets = parse_sheets(&workbook_xml)?;
    let targets = parse_relationship_targets(&rels_xml)?;
    let names: Vec<String> = sheets.iter().map(|sheet| sheet.name.clone()).collect();
    if let Some(unknown) = visibilities.unknown(&names) {
        return Err(Error::SheetNotFound {
            sheet: unknown.to_string(),
            part: WORKBOOK_PART.to_string(),
        });
    }
    let states = SheetStates::read(&workbook_xml, &sheets, visibilities)?;
    let mut plans: HashMap<String, SheetPlan> = HashMap::new();
    for (i, sheet) in sheets.iter().enumerate() {
        let mut plan = SheetPlan::default();
        if let Some(sheet_visibility) = visibilities.sheets.get(&sheet.name) {
            plan.rows = Some(&sheet_visibility.rows);
            plan.columns = Some(&sheet_visibility.columns);
        }
        if states.deselected.contains(&i) {
            plan.tab_selected = Some(false);
        } else if states.selected == Some(i) {
            plan.tab_selected = Some(true);
        }
        if plan.is_empty() {
            continue;
        }
        let target = targets.get(&sheet.rel_id).with_context(|| {
            format!(
                "Relationship '{}' not found for sheet '{}'",
                sheet.rel_id, sheet.name
            )
        })?;
        plans.insert(workbook_part_name(target), plan);
    }
    for i in 0..archive.len() {
        let name = archive.by_index_raw(i)?.name()?.into_owned();
        if name == WORKBOOK_PART && states.changes {
            zip_writer.start_file(&name, part_options())?;
            zip_writer.write_all(states.rewrite(&workbook_xml)?.as_bytes())?;
        } else if let Some(plan) = plans.get(&name) {
            zip_writer.start_file(&name, part_options())?;
            rewrite_sheet(BufReader::new(archive.by_index(i)?), &mut *zip_writer, plan)
                .with_context(|| format!("Failed to rewrite {name}"))?;
        } else {
            zip_writer.raw_copy_file(archive.by_index_raw(i)?)?;
        }
    }
    Ok(())
}

/// What the workbook says about its sheets and what they become.
struct SheetStates {
    /// The `state` each sheet has at the end, by position (`None`: visible).
    after: Vec<Option<String>>,
    /// Whether any sheet changes state or the workbook changes its active tab.
    changes: bool,
    /// The sheets that were hidden while they were the selected tab.
    deselected: Vec<usize>,
    /// The tab that is selected after the pass, when it is not the one that was.
    selected: Option<usize>,
    /// The `activeTab` after the pass, when it changes.
    active_tab: Option<usize>,
    /// The `firstSheet` after the pass, when it changes.
    first_sheet: Option<usize>,
}

impl SheetStates {
    fn read(xml: &str, sheets: &[SheetEntry], visibilities: &Visibilities) -> Result<Self> {
        let mut reader = Reader::from_str(xml);
        let mut before: Vec<Option<String>> = Vec::new();
        let (mut active, mut first) = (0usize, 0usize);
        loop {
            match reader.read_event()? {
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "sheet" => {
                    before.push(current_state(&e)?);
                }
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "workbookView" => {
                    active = attribute(&e, "activeTab")?
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0);
                    first = attribute(&e, "firstSheet")?
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0);
                }
                Event::Eof => break,
                _ => {}
            }
        }
        let mut after = before.clone();
        let mut changes = false;
        for (i, sheet) in sheets.iter().enumerate() {
            let Some(&hidden) = visibilities.states.get(&sheet.name) else {
                continue;
            };
            let new = match (hidden, &before[i]) {
                // Hiding what is very hidden leaves it so.
                (true, Some(state)) => Some(state.clone()),
                (true, None) => Some("hidden".to_string()),
                (false, _) => None,
            };
            changes |= new != before[i];
            after[i] = new;
        }
        let first_visible = after
            .iter()
            .position(Option::is_none)
            .ok_or(Error::NoVisibleSheet)?;
        let hidden_now = |i: usize| after.get(i).is_some_and(Option::is_some);
        let active_hidden = hidden_now(active);
        let deselected = if active_hidden && before.get(active).is_some_and(Option::is_none) {
            vec![active]
        } else {
            Vec::new()
        };
        let (active_tab, selected) = if active_hidden {
            changes = true;
            (Some(first_visible), Some(first_visible))
        } else {
            (None, None)
        };
        let first_sheet = hidden_now(first).then_some(first_visible);
        changes |= first_sheet.is_some();
        Ok(Self {
            after,
            changes,
            deselected,
            selected,
            active_tab,
            first_sheet,
        })
    }

    /// `workbook.xml` with the states of the sheets and the tab the workbook opens on.
    fn rewrite(&self, xml: &str) -> Result<String> {
        let mut reader = Reader::from_str(xml);
        let mut writer = Writer::new(std::io::Cursor::new(Vec::new()));
        let mut index = 0usize;
        loop {
            match reader.read_event()? {
                Event::Start(e) if e.name().as_ref() == "sheet" => {
                    writer.write_event(Event::Start(self.with_state(&e, index)?))?;
                    index += 1;
                }
                Event::Empty(e) if e.name().as_ref() == "sheet" => {
                    writer.write_event(Event::Empty(self.with_state(&e, index)?))?;
                    index += 1;
                }
                Event::Start(e) if e.name().as_ref() == "workbookView" => {
                    writer.write_event(Event::Start(self.with_tabs(&e)?))?;
                }
                Event::Empty(e) if e.name().as_ref() == "workbookView" => {
                    writer.write_event(Event::Empty(self.with_tabs(&e)?))?;
                }
                Event::Eof => break,
                event => writer.write_event(event)?,
            }
        }
        Ok(String::from_utf8(writer.into_inner().into_inner())?)
    }

    fn with_state(&self, sheet: &BytesStart, index: usize) -> Result<BytesStart<'static>> {
        rebuilt(sheet, "state", self.after[index].as_deref())
    }

    fn with_tabs(&self, view: &BytesStart) -> Result<BytesStart<'static>> {
        let mut view = view.to_owned();
        if let Some(active) = self.active_tab {
            view = rebuilt(&view, "activeTab", Some(&active.to_string()))?;
        }
        if let Some(first) = self.first_sheet {
            view = rebuilt(&view, "firstSheet", Some(&first.to_string()))?;
        }
        Ok(view)
    }
}

/// `element` with the attribute `key` set to `value`, or without it when there is none.
fn rebuilt(element: &BytesStart, key: &str, value: Option<&str>) -> Result<BytesStart<'static>> {
    let mut out = BytesStart::new(element.name().as_ref().to_string());
    for attr in element.attributes() {
        let attr = attr?;
        if attr.key.as_ref() != key {
            out.push_attribute((attr.key.as_ref(), &*attr_value(&attr)?));
        }
    }
    if let Some(value) = value {
        out.push_attribute((key, value));
    }
    Ok(out)
}

fn is_true(value: &str) -> bool {
    matches!(value, "1" | "true")
}

/// A `<row>` with the visibility asked: no change where `hidden` is `None`.
fn row_with(row: &BytesStart, hidden: Option<bool>) -> Result<BytesStart<'static>> {
    match hidden {
        None => Ok(row.to_owned()),
        Some(true) => rebuilt(row, "hidden", Some("1")),
        Some(false) => rebuilt(row, "hidden", None),
    }
}

/// A row that the sheet did not have, hidden.
fn missing_row(row: u32) -> BytesStart<'static> {
    let mut element = BytesStart::new("row");
    element.push_attribute(("r", (row + 1).to_string().as_str()));
    element.push_attribute(("hidden", "1"));
    element
}

/// The hidden rows that the sheet does not have yet, written as the rows go by.
struct MissingRows<'a> {
    hidden: Option<&'a Intervals>,
    /// The first row that has not been decided: the rows before it are written or exist.
    next: u32,
}

impl MissingRows<'_> {
    /// Write the hidden rows before `row` that no `<row>` has.
    fn write_before<W: Write>(&mut self, writer: &mut Writer<W>, row: u32) -> Result<()> {
        let Some(intervals) = self.hidden else {
            return Ok(());
        };
        while let Some(hidden) = intervals
            .next_hidden(self.next)
            .filter(|&hidden| hidden < row && hidden < MAX_ROW)
        {
            writer.write_event(Event::Empty(missing_row(hidden)))?;
            self.next = hidden + 1;
        }
        self.next = self.next.max(row);
        Ok(())
    }
}

/// One `<col>` as read: its columns, the visibility and the attributes the pass keeps as they are.
#[derive(Clone, PartialEq)]
struct Col {
    min: u32,
    max: u32,
    hidden: bool,
    rest: Vec<(String, String)>,
}

impl Col {
    fn read(element: &BytesStart) -> Result<Option<Self>> {
        let (mut min, mut max, mut hidden, mut rest) = (None, None, false, Vec::new());
        for attr in element.attributes() {
            let attr = attr?;
            let value = attr_value(&attr)?.into_owned();
            match attr.key.as_ref() {
                "min" => min = value.parse::<u32>().ok(),
                "max" => max = value.parse::<u32>().ok(),
                "hidden" => hidden = is_true(&value),
                key => rest.push((key.to_string(), value)),
            }
        }
        Ok(match (min, max) {
            (Some(min), Some(max)) if min >= 1 && max >= min => Some(Self {
                min: min - 1,
                max: max - 1,
                hidden,
                rest,
            }),
            _ => None,
        })
    }

    fn element(&self) -> BytesStart<'static> {
        let mut element = BytesStart::new("col");
        element.push_attribute(("min", (self.min + 1).to_string().as_str()));
        element.push_attribute(("max", (self.max + 1).to_string().as_str()));
        for (key, value) in &self.rest {
            element.push_attribute((key.as_str(), value.as_str()));
        }
        if self.hidden {
            element.push_attribute(("hidden", "1"));
        }
        element
    }

    /// Whether `other` goes on from this one with the same look.
    fn continues_into(&self, other: &Self) -> bool {
        self.max + 1 == other.min && self.hidden == other.hidden && self.rest == other.rest
    }
}

/// The `<col>` elements after the visibility asked: the declared columns cut where the state
/// changes, and a column of its own for each hidden run that no `<col>` declared.
fn columns_after(declared: &[Col], asked: &Intervals, default_width: &str) -> Vec<Col> {
    let mut cuts: Vec<u32> = asked.boundaries().collect();
    for col in declared {
        cuts.push(col.min);
        cuts.push(col.max + 1);
    }
    cuts.sort_unstable();
    cuts.dedup();
    let mut out: Vec<Col> = Vec::new();
    for window in cuts.windows(2) {
        let (start, end) = (window[0], window[1] - 1);
        if start > MAX_COL {
            break;
        }
        let end = end.min(MAX_COL);
        let base = declared.iter().find(|c| c.min <= start && start <= c.max);
        let asked_hidden = asked.state_at(start);
        let piece = match (base, asked_hidden) {
            (Some(base), state) => Col {
                min: start,
                max: end,
                hidden: state.unwrap_or(base.hidden),
                rest: base.rest.clone(),
            },
            (None, Some(true)) => Col {
                min: start,
                max: end,
                hidden: true,
                rest: vec![
                    ("width".to_string(), default_width.to_string()),
                    ("customWidth".to_string(), "1".to_string()),
                ],
            },
            (None, _) => continue,
        };
        match out.last_mut() {
            Some(last) if last.continues_into(&piece) => last.max = piece.max,
            _ => out.push(piece),
        }
    }
    out
}

fn write_cols<W: Write>(writer: &mut Writer<W>, cols: &[Col]) -> Result<()> {
    if cols.is_empty() {
        return Ok(());
    }
    writer.write_event(Event::Start(BytesStart::new("cols")))?;
    for col in cols {
        writer.write_event(Event::Empty(col.element()))?;
    }
    writer.write_event(Event::End(BytesEnd::new("cols")))?;
    Ok(())
}

/// What the pass keeps while it reads the worksheet.
struct SheetState<'a> {
    plan: &'a SheetPlan<'a>,
    rows: Option<&'a Intervals>,
    columns: Option<&'a Intervals>,
    missing: MissingRows<'a>,
    default_width: String,
    cols_written: bool,
    last_row: Option<u32>,
}

impl SheetState<'_> {
    /// The `<cols>` of a sheet that had none, before its data.
    fn write_new_cols<W: Write>(&mut self, writer: &mut Writer<W>) -> Result<()> {
        if let (false, Some(asked)) = (self.cols_written, self.columns) {
            write_cols(writer, &columns_after(&[], asked, &self.default_width))?;
        }
        self.cols_written = true;
        Ok(())
    }

    /// One element that opens (`empty` false) or is empty.
    fn element<R: BufRead, W: Write>(
        &mut self,
        reader: &mut Reader<R>,
        writer: &mut Writer<W>,
        e: BytesStart<'static>,
        empty: bool,
    ) -> Result<()> {
        let write = |writer: &mut Writer<W>, e: BytesStart<'static>| {
            writer.write_event(if empty {
                Event::Empty(e)
            } else {
                Event::Start(e)
            })
        };
        match e.name().as_ref() {
            "sheetFormatPr" => {
                if let Some(width) = attribute(&e, "defaultColWidth")? {
                    self.default_width = width;
                }
                write(writer, e)?;
            }
            "cols" if !empty && self.columns.is_some() => {
                let mut declared = Vec::new();
                let mut buf = Vec::new();
                loop {
                    buf.clear();
                    match reader.read_event_into(&mut buf)? {
                        Event::Start(c) | Event::Empty(c) if c.name().as_ref() == "col" => {
                            if let Some(col) = Col::read(&c)? {
                                declared.push(col);
                            }
                        }
                        Event::End(c) if c.name().as_ref() == "cols" => break,
                        Event::Eof => return Err(Error::malformed("unexpected end inside <cols>")),
                        _ => {}
                    }
                }
                declared.sort_by_key(|col| col.min);
                let asked = self.columns.expect("matched above");
                write_cols(
                    writer,
                    &columns_after(&declared, asked, &self.default_width),
                )?;
                self.cols_written = true;
            }
            "sheetData" => {
                self.write_new_cols(writer)?;
                if empty {
                    writer.write_event(Event::Start(e))?;
                    self.missing.write_before(writer, MAX_ROW)?;
                    writer.write_event(Event::End(BytesEnd::new("sheetData")))?;
                } else {
                    writer.write_event(Event::Start(e))?;
                }
            }
            "row" => {
                let number = attribute(&e, "r")?
                    .and_then(|r| r.parse::<u32>().ok())
                    .map_or_else(
                        || self.last_row.map_or(0, |r| r + 1),
                        |r| r.saturating_sub(1),
                    );
                self.last_row = Some(number);
                self.missing.write_before(writer, number)?;
                self.missing.next = number + 1;
                let hidden = self.rows.and_then(|rows| rows.state_at(number));
                write(writer, row_with(&e, hidden)?)?;
            }
            "sheetView" if self.plan.tab_selected.is_some() => {
                let value = self.plan.tab_selected.expect("checked").then_some("1");
                write(writer, rebuilt(&e, "tabSelected", value)?)?;
            }
            _ => write(writer, e)?,
        }
        Ok(())
    }
}

fn rewrite_sheet<R: BufRead, W: Write>(input: R, output: W, plan: &SheetPlan) -> Result<()> {
    let mut reader = Reader::from_reader(input);
    let mut writer = Writer::new(BufWriter::new(output));
    let rows = plan.rows.filter(|rows| !rows.is_empty());
    let columns = plan.columns.filter(|columns| !columns.is_empty());
    let mut state = SheetState {
        plan,
        rows,
        columns,
        missing: MissingRows {
            hidden: rows,
            next: 0,
        },
        default_width: DEFAULT_COLUMN_WIDTH.to_string(),
        cols_written: false,
        last_row: None,
    };
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => state.element(&mut reader, &mut writer, e.into_owned(), false)?,
            Event::Empty(e) => state.element(&mut reader, &mut writer, e.into_owned(), true)?,
            Event::End(e) if e.name().as_ref() == "sheetData" => {
                state.missing.write_before(&mut writer, MAX_ROW)?;
                writer.write_event(Event::End(e))?;
            }
            Event::Eof => break,
            event => writer.write_event(event)?,
        }
    }
    writer.into_inner().flush()?;
    Ok(())
}

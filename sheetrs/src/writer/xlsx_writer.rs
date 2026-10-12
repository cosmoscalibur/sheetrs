// ! XLSX writer functionality for removing sheets and named ranges, and for
// ! converting formulas to values

use super::ValueScope;
use crate::reader::parser_utils::{attr_value, parse_cell_range, parse_cell_ref, push_general_ref};
use crate::reader::xlsx_parser::{get_xlsx_sheet_path, translate_shared_formula};
use anyhow::{Context, Result};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read, Seek, Write};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

/// Part name of the calculation chain, which lists every formula cell.
const CALC_CHAIN_PART: &str = "xl/calcChain.xml";

/// Struct used to define modifications to be applied to a workbook
#[derive(Debug, Default)]
pub struct WorkbookModifications {
    pub remove_sheets: Option<HashSet<String>>,
    pub remove_named_ranges: Option<HashSet<String>>,
    /// Replace formulas by their cached values in these cells.
    pub convert_to_values: Vec<ValueScope>,
}

/// Modify an XLSX file by applying specified modifications
pub fn modify_workbook_xlsx<R: Read + Seek, W: Write + Seek>(
    mut input: R,
    output: W,
    modifications: &WorkbookModifications,
) -> Result<()> {
    let mut archive = ZipArchive::new(&mut input)?;

    // Create output ZIP
    let mut zip_writer = ZipWriter::new(output);

    // Read workbook.xml needed for IDs and cleanup
    let workbook_xml = read_file_from_zip(&mut archive, "xl/workbook.xml")?;
    let sheet_info = parse_sheet_info(&workbook_xml)?;

    // Identify sheet IDs to remove
    let sheet_ids_to_remove: Vec<usize> = if let Some(sheets) = &modifications.remove_sheets {
        sheet_info
            .iter()
            .filter(|(name, _)| sheets.contains(name))
            .map(|(_, id)| *id)
            .collect()
    } else {
        Vec::new()
    };

    // Resolve the worksheet part of every sheet to convert to values
    let mut value_scopes: HashMap<String, Vec<&ValueScope>> = HashMap::new();
    for scope in &modifications.convert_to_values {
        if modifications
            .remove_sheets
            .as_ref()
            .is_some_and(|sheets| sheets.contains(&scope.sheet))
        {
            continue;
        }
        let path = get_xlsx_sheet_path(&mut archive, &scope.sheet)?;
        value_scopes.entry(path).or_default().push(scope);
    }
    // Excel rebuilds the calculation chain when it is missing, but reports a
    // corrupt file when it lists cells that no longer hold formulas.
    let drop_calc_chain = !value_scopes.is_empty();

    // Iterate through all files in the archive
    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let name = file.name().to_string();

        // Check if we should skip this file (e.g. removed sheet xml)
        let should_skip = sheet_ids_to_remove.iter().any(|id| {
            name == format!("xl/worksheets/sheet{id}.xml")
                || name == format!("xl/worksheets/_rels/sheet{id}.xml.rels")
        });

        if should_skip || (drop_calc_chain && name == CALC_CHAIN_PART) {
            continue;
        }

        // Processing logic
        if name == "xl/workbook.xml" {
            let mut content = workbook_xml.clone(); // Use cached content

            // 1. Remove sheets
            if let Some(sheets) = &modifications.remove_sheets
                && !sheets.is_empty()
            {
                content = remove_sheets_from_workbook_xml(&content, sheets)?;
            }

            // 2. Remove named ranges
            if let Some(ranges) = &modifications.remove_named_ranges
                && !ranges.is_empty()
            {
                content = remove_named_ranges_from_workbook_xml(&content, ranges)?;
            }

            zip_writer.start_file(&name, SimpleFileOptions::default())?;
            zip_writer.write_all(content.as_bytes())?;
        } else if name == "[Content_Types].xml"
            && (!sheet_ids_to_remove.is_empty() || drop_calc_chain)
        {
            let mut content = String::new();
            file.read_to_string(&mut content)?;
            let modified_content = remove_content_type_overrides(&content, |part_name| {
                (drop_calc_chain && part_name.strip_prefix('/') == Some(CALC_CHAIN_PART))
                    || sheet_ids_to_remove
                        .iter()
                        .any(|id| part_name == format!("/xl/worksheets/sheet{id}.xml"))
            })?;
            zip_writer.start_file(&name, SimpleFileOptions::default())?;
            zip_writer.write_all(modified_content.as_bytes())?;
        } else if name == "xl/_rels/workbook.xml.rels"
            && (!sheet_ids_to_remove.is_empty() || drop_calc_chain)
        {
            let mut content = String::new();
            file.read_to_string(&mut content)?;
            let modified_content = remove_relationships(&content, |target| {
                (drop_calc_chain && is_calc_chain_target(target))
                    || sheet_ids_to_remove
                        .iter()
                        .any(|id| target == format!("worksheets/sheet{id}.xml"))
            })?;
            zip_writer.start_file(&name, SimpleFileOptions::default())?;
            zip_writer.write_all(modified_content.as_bytes())?;
        } else if let Some(scopes) = value_scopes.get(&name) {
            let mut content = String::new();
            file.read_to_string(&mut content)?;
            let modified_content = convert_formulas_to_values(&content, scopes)
                .with_context(|| format!("Failed to convert formulas to values in {name}"))?;
            zip_writer.start_file(&name, SimpleFileOptions::default())?;
            zip_writer.write_all(modified_content.as_bytes())?;
        } else {
            // Copy file as is
            zip_writer.start_file(&name, SimpleFileOptions::default())?;
            let mut buffer = Vec::new();
            file.read_to_end(&mut buffer)?;
            zip_writer.write_all(&buffer)?;
        }
    }

    zip_writer.finish()?;
    Ok(())
}

// Helper functions

fn read_file_from_zip<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    filename: &str,
) -> Result<String> {
    let mut file = archive.by_name(filename)?;
    let mut content = String::new();
    file.read_to_string(&mut content)?;
    Ok(content)
}

fn parse_sheet_info(workbook_xml: &str) -> Result<Vec<(String, usize)>> {
    let mut reader = Reader::from_str(workbook_xml);
    let mut buf = Vec::new();
    let mut sheets = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.name().as_ref() == "sheet" => {
                let mut name = String::new();
                let mut sheet_id = 0;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        "name" => {
                            name = attr_value(&attr)?.into_owned();
                        }
                        "sheetId" => {
                            sheet_id = attr_value(&attr)?.into_owned().parse()?;
                        }
                        _ => {}
                    }
                }

                sheets.push((name, sheet_id));
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("Error parsing XML: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(sheets)
}

fn remove_sheets_from_workbook_xml(
    xml: &str,
    sheets_to_remove: &HashSet<String>,
) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();
    let mut skip_current_sheet = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == "sheet" => {
                let mut sheet_name = String::new();
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == "name" {
                        sheet_name = attr_value(&attr)?.into_owned();
                        break;
                    }
                }

                if sheets_to_remove.contains(&sheet_name) {
                    skip_current_sheet = true;
                } else {
                    writer.write_event(Event::Start(e))?;
                }
            }
            Ok(Event::Empty(e)) if e.name().as_ref() == "sheet" => {
                let mut sheet_name = String::new();
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == "name" {
                        sheet_name = attr_value(&attr)?.into_owned();
                        break;
                    }
                }

                if !sheets_to_remove.contains(&sheet_name) {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Ok(Event::End(e)) if e.name().as_ref() == "sheet" => {
                if skip_current_sheet {
                    skip_current_sheet = false;
                } else {
                    writer.write_event(Event::End(e))?;
                }
            }
            Ok(Event::Eof) => break,
            Ok(e) => {
                if !skip_current_sheet {
                    writer.write_event(e)?;
                }
            }
            Err(e) => return Err(anyhow::anyhow!("Error parsing XML: {e}")),
        }
        buf.clear();
    }

    let result = writer.into_inner().into_inner();
    Ok(String::from_utf8(result)?)
}

/// Drop the `<Override>` entries whose `PartName` matches `remove`.
fn remove_content_type_overrides(xml: &str, remove: impl Fn(&str) -> bool) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) if e.name().as_ref() == "Override" => {
                let mut part_name = String::new();
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == "PartName" {
                        part_name = attr_value(&attr)?.into_owned();
                        break;
                    }
                }

                if !remove(&part_name) {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Ok(Event::Eof) => break,
            Ok(e) => writer.write_event(e)?,
            Err(e) => return Err(anyhow::anyhow!("Error parsing XML: {e}")),
        }
        buf.clear();
    }

    let result = writer.into_inner().into_inner();
    Ok(String::from_utf8(result)?)
}

/// Drop the `<Relationship>` entries whose `Target` matches `remove`.
fn remove_relationships(xml: &str, remove: impl Fn(&str) -> bool) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) if e.name().as_ref() == "Relationship" => {
                let mut target = String::new();
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == "Target" {
                        target = attr_value(&attr)?.into_owned();
                        break;
                    }
                }

                if !remove(&target) {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Ok(Event::Eof) => break,
            Ok(e) => writer.write_event(e)?,
            Err(e) => return Err(anyhow::anyhow!("Error parsing XML: {e}")),
        }
        buf.clear();
    }

    let result = writer.into_inner().into_inner();
    Ok(String::from_utf8(result)?)
}

fn remove_named_ranges_from_workbook_xml(
    xml: &str,
    ranges_to_remove: &HashSet<String>,
) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();
    let mut skip_current_range = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == "definedName" => {
                let mut range_name = String::new();
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == "name" {
                        range_name = attr_value(&attr)?.into_owned();
                        break;
                    }
                }

                if ranges_to_remove.contains(&range_name) {
                    skip_current_range = true;
                } else {
                    writer.write_event(Event::Start(e))?;
                }
            }
            Ok(Event::Empty(e)) if e.name().as_ref() == "definedName" => {
                let mut range_name = String::new();
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == "name" {
                        range_name = attr_value(&attr)?.into_owned();
                        break;
                    }
                }

                if !ranges_to_remove.contains(&range_name) {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Ok(Event::End(e)) if e.name().as_ref() == "definedName" => {
                if skip_current_range {
                    skip_current_range = false;
                } else {
                    writer.write_event(Event::End(e))?;
                }
            }
            Ok(Event::Eof) => break,
            Ok(e) => {
                if !skip_current_range {
                    writer.write_event(e)?;
                }
            }
            Err(e) => return Err(anyhow::anyhow!("Error parsing XML: {e}")),
        }
        buf.clear();
    }

    let result = writer.into_inner().into_inner();
    Ok(String::from_utf8(result)?)
}

fn is_calc_chain_target(target: &str) -> bool {
    target == "calcChain.xml" || target.strip_prefix('/') == Some(CALC_CHAIN_PART)
}

/// `<f>` element of a cell.
struct CellFormula {
    /// The `t` attribute: `shared`, `array`, `dataTable` or none (normal).
    kind: Option<String>,
    shared_index: Option<u32>,
    /// The `ref` attribute: cells of a shared group, array or data table.
    range: Option<(u32, u32, u32, u32)>,
    /// Formula text; empty for the cells of a shared group but its first.
    text: String,
}

impl CellFormula {
    fn is_shared(&self) -> bool {
        self.kind.as_deref() == Some("shared")
    }
}

/// Formula text of a shared group, held by its first cell.
struct SharedFormula {
    row: u32,
    col: u32,
    range: Option<(u32, u32, u32, u32)>,
    text: String,
    /// Whether that first cell is converted to its value, which leaves the
    /// rest of the group without formula text.
    converted: bool,
}

/// Position of the `<c>` elements, which may omit their `r` attribute.
#[derive(Default)]
struct CellCursor {
    row: u32,
    next_row: u32,
    next_col: u32,
}

impl CellCursor {
    fn start_row(&mut self, row: &BytesStart) -> Result<()> {
        self.row = match attribute(row, "r")? {
            Some(r) => r.parse::<u32>()?.saturating_sub(1),
            None => self.next_row,
        };
        self.next_row = self.row + 1;
        self.next_col = 0;
        Ok(())
    }

    fn next_cell(&mut self, cell: &BytesStart) -> Result<(u32, u32)> {
        let (row, col) = attribute(cell, "r")?
            .and_then(|r| parse_cell_ref(&r))
            .unwrap_or((self.row, self.next_col));
        self.next_col = col + 1;
        Ok((row, col))
    }
}

fn attribute(element: &BytesStart, key: &str) -> Result<Option<String>> {
    for attr in element.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == key {
            return Ok(Some(attr_value(&attr)?.into_owned()));
        }
    }
    Ok(None)
}

/// Read the events of a `<c>` element up to, not including, its end tag.
fn read_cell_body<'a>(reader: &mut Reader<&'a [u8]>) -> Result<Vec<Event<'a>>> {
    let mut body = Vec::new();
    loop {
        match reader.read_event()? {
            Event::End(e) if e.name().as_ref() == "c" => return Ok(body),
            Event::Eof => anyhow::bail!("unexpected end of worksheet inside a cell"),
            event => body.push(event),
        }
    }
}

/// The formula of a cell and whether it holds a cached value.
fn inspect_cell(body: &[Event]) -> Result<(Option<CellFormula>, bool)> {
    let mut formula: Option<CellFormula> = None;
    let mut has_value = false;
    let mut in_formula = false;
    for event in body {
        match event {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "f" => {
                formula = Some(CellFormula {
                    kind: attribute(e, "t")?,
                    shared_index: attribute(e, "si")?.map(|si| si.parse()).transpose()?,
                    range: attribute(e, "ref")?.and_then(|r| parse_range(&r)),
                    text: String::new(),
                });
                in_formula = matches!(event, Event::Start(_));
            }
            Event::End(e) if e.name().as_ref() == "f" => in_formula = false,
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "v" => has_value = true,
            Event::Text(text) if in_formula => {
                if let Some(formula) = &mut formula {
                    formula.text.push_str(text);
                }
            }
            Event::CData(text) if in_formula => {
                if let Some(formula) = &mut formula {
                    formula.text.push_str(text);
                }
            }
            Event::GeneralRef(reference) if in_formula => {
                if let Some(formula) = &mut formula {
                    push_general_ref(&mut formula.text, reference)?;
                }
            }
            _ => {}
        }
    }
    Ok((formula, has_value))
}

/// Parse a `ref` attribute, which names a single cell when it spans one.
fn parse_range(reference: &str) -> Option<(u32, u32, u32, u32)> {
    parse_cell_range(reference)
        .or_else(|| parse_cell_ref(reference).map(|(row, col)| (row, col, row, col)))
}

/// Whether a formula cell is replaced by its cached value.
///
/// Cells without a cached value keep their formula. Array formulas and data
/// tables are converted only when one scope covers their whole range, since
/// their value cells hold no formula of their own.
fn converts_to_value(
    scopes: &[&ValueScope],
    row: u32,
    col: u32,
    formula: &CellFormula,
    has_value: bool,
) -> bool {
    if !has_value {
        return false;
    }
    match formula.kind.as_deref() {
        Some("array" | "dataTable") => {
            let (first_row, first_col, last_row, last_col) =
                formula.range.unwrap_or((row, col, row, col));
            // Scopes are whole sheets, column bands or row bands, so holding
            // two opposite corners means holding the whole range.
            scopes.iter().any(|scope| {
                scope.contains(first_row, first_col) && scope.contains(last_row, last_col)
            })
        }
        _ => scopes.iter().any(|scope| scope.contains(row, col)),
    }
}

/// Collect the formula text of every shared group, keyed by `si`.
fn collect_shared_formulas(
    xml: &str,
    scopes: &[&ValueScope],
) -> Result<HashMap<u32, Vec<SharedFormula>>> {
    let mut reader = Reader::from_str(xml);
    let mut cursor = CellCursor::default();
    let mut shared: HashMap<u32, Vec<SharedFormula>> = HashMap::new();
    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "row" => {
                cursor.start_row(&e)?;
            }
            Event::Start(e) if e.name().as_ref() == "c" => {
                let (row, col) = cursor.next_cell(&e)?;
                let body = read_cell_body(&mut reader)?;
                if let (Some(formula), has_value) = inspect_cell(&body)?
                    && formula.is_shared()
                    && !formula.text.is_empty()
                    && let Some(si) = formula.shared_index
                {
                    let converted = converts_to_value(scopes, row, col, &formula, has_value);
                    shared.entry(si).or_default().push(SharedFormula {
                        row,
                        col,
                        range: formula.range,
                        text: formula.text,
                        converted,
                    });
                }
            }
            Event::Empty(e) if e.name().as_ref() == "c" => {
                cursor.next_cell(&e)?;
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(shared)
}

/// Find the group a shared formula cell belongs to: the one whose range holds
/// the cell, else the last one with its index.
fn find_shared_formula(groups: &[SharedFormula], row: u32, col: u32) -> Option<&SharedFormula> {
    groups
        .iter()
        .find(|group| {
            group
                .range
                .is_some_and(|(first_row, first_col, last_row, last_col)| {
                    (first_row..=last_row).contains(&row) && (first_col..=last_col).contains(&col)
                })
        })
        .or(groups.last())
}

/// Replace the formulas of the cells in `scopes` by their cached values.
///
/// A cell of a shared group whose first cell is converted, but that is not
/// converted itself, gets its own copy of the translated formula.
fn convert_formulas_to_values(xml: &str, scopes: &[&ValueScope]) -> Result<String> {
    let shared = collect_shared_formulas(xml, scopes)?;
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::with_capacity(xml.len())));
    let mut cursor = CellCursor::default();

    loop {
        match reader.read_event()? {
            Event::Start(e) if e.name().as_ref() == "row" => {
                cursor.start_row(&e)?;
                writer.write_event(Event::Start(e))?;
            }
            Event::Empty(e) if e.name().as_ref() == "row" => {
                cursor.start_row(&e)?;
                writer.write_event(Event::Empty(e))?;
            }
            Event::Start(e) if e.name().as_ref() == "c" => {
                let (row, col) = cursor.next_cell(&e)?;
                let body = read_cell_body(&mut reader)?;
                let (formula, has_value) = inspect_cell(&body)?;
                match formula {
                    Some(formula) if converts_to_value(scopes, row, col, &formula, has_value) => {
                        write_value_cell(&mut writer, &e, &body)?;
                    }
                    Some(formula) if formula.is_shared() && formula.text.is_empty() => {
                        let unshared = formula
                            .shared_index
                            .and_then(|si| shared.get(&si))
                            .and_then(|groups| find_shared_formula(groups, row, col))
                            .filter(|group| group.converted)
                            .and_then(|group| {
                                Some(translate_shared_formula(
                                    &group.text,
                                    row.checked_signed_diff(group.row)?,
                                    col.checked_signed_diff(group.col)?,
                                ))
                            });
                        match unshared {
                            Some(text) => write_unshared_cell(&mut writer, e, &body, &text)?,
                            None => write_cell(&mut writer, e, &body)?,
                        }
                    }
                    _ => write_cell(&mut writer, e, &body)?,
                }
            }
            Event::Empty(e) if e.name().as_ref() == "c" => {
                cursor.next_cell(&e)?;
                writer.write_event(Event::Empty(e))?;
            }
            Event::Eof => break,
            event => writer.write_event(event)?,
        }
    }

    Ok(String::from_utf8(writer.into_inner().into_inner())?)
}

fn write_cell<W: Write>(writer: &mut Writer<W>, start: BytesStart, body: &[Event]) -> Result<()> {
    let end = start.to_end().into_owned();
    writer.write_event(Event::Start(start))?;
    for event in body {
        writer.write_event(event.borrow())?;
    }
    writer.write_event(Event::End(end))?;
    Ok(())
}

/// Write a cell without its `<f>` element. A formula string result
/// (`t="str"`) becomes an inline string, since `str` marks formula cells.
fn write_value_cell<W: Write>(
    writer: &mut Writer<W>,
    start: &BytesStart,
    body: &[Event],
) -> Result<()> {
    let inline_string = attribute(start, "t")?.as_deref() == Some("str");
    let mut cell = start.clone();
    cell.clear_attributes();
    for attr in start.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            // Cell metadata marks dynamic array formulas.
            "cm" => {}
            "t" if inline_string => cell.push_attribute(("t", "inlineStr")),
            _ => cell.push_attribute(attr),
        }
    }
    let mut text = BytesStart::new("t");
    text.push_attribute(("xml:space", "preserve"));

    let end = cell.to_end().into_owned();
    writer.write_event(Event::Start(cell))?;
    let mut in_formula = false;
    for event in body {
        match event {
            Event::Start(e) if e.name().as_ref() == "f" => in_formula = true,
            Event::End(e) if e.name().as_ref() == "f" => in_formula = false,
            Event::Empty(e) if e.name().as_ref() == "f" => {}
            _ if in_formula => {}
            Event::Start(e) if inline_string && e.name().as_ref() == "v" => {
                writer.write_event(Event::Start(BytesStart::new("is")))?;
                writer.write_event(Event::Start(text.borrow()))?;
            }
            Event::End(e) if inline_string && e.name().as_ref() == "v" => {
                writer.write_event(Event::End(BytesEnd::new("t")))?;
                writer.write_event(Event::End(BytesEnd::new("is")))?;
            }
            Event::Empty(e) if inline_string && e.name().as_ref() == "v" => {
                writer.write_event(Event::Start(BytesStart::new("is")))?;
                writer.write_event(Event::Empty(text.borrow()))?;
                writer.write_event(Event::End(BytesEnd::new("is")))?;
            }
            event => writer.write_event(event.borrow())?,
        }
    }
    writer.write_event(Event::End(end))?;
    Ok(())
}

/// Write a cell of a shared group with `formula` as its own formula text.
fn write_unshared_cell<W: Write>(
    writer: &mut Writer<W>,
    start: BytesStart,
    body: &[Event],
    formula: &str,
) -> Result<()> {
    let end = start.to_end().into_owned();
    writer.write_event(Event::Start(start))?;
    let mut in_formula = false;
    for event in body {
        match event {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "f" => {
                writer.write_event(Event::Start(BytesStart::new("f")))?;
                writer.write_event(Event::Text(BytesText::new(formula)))?;
                writer.write_event(Event::End(BytesEnd::new("f")))?;
                in_formula = matches!(event, Event::Start(_));
            }
            Event::End(e) if e.name().as_ref() == "f" => in_formula = false,
            _ if in_formula => {}
            event => writer.write_event(event.borrow())?,
        }
    }
    writer.write_event(Event::End(end))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    const WORKBOOK_XML: &str = r#"<workbook><sheets><sheet name="R&amp;D" sheetId="1"/><sheet name="Data" sheetId="2"/></sheets><definedNames><definedName name="A&amp;B">Data!$A$1</definedName><definedName name="Keep">Data!$B$1</definedName></definedNames></workbook>"#;

    #[test]
    fn test_parse_sheet_info_unescapes_names() {
        let sheets = parse_sheet_info(WORKBOOK_XML).unwrap();
        assert_eq!(
            sheets,
            vec![("R&D".to_string(), 1), ("Data".to_string(), 2)]
        );
    }

    #[test]
    fn test_remove_sheet_with_escaped_name() {
        let remove = HashSet::from(["R&D".to_string()]);
        let xml = remove_sheets_from_workbook_xml(WORKBOOK_XML, &remove).unwrap();
        assert!(!xml.contains("R&amp;D"));
        assert!(xml.contains(r#"<sheet name="Data" sheetId="2"/>"#));
    }

    #[test]
    fn test_remove_named_range_with_escaped_name() {
        let remove = HashSet::from(["A&B".to_string()]);
        let xml = remove_named_ranges_from_workbook_xml(WORKBOOK_XML, &remove).unwrap();
        assert!(!xml.contains("A&amp;B"));
        assert!(xml.contains(r#"<definedName name="Keep">Data!$B$1</definedName>"#));
    }
}

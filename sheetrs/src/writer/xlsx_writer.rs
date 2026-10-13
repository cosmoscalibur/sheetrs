//! XLSX writer functionality for removing sheets and named ranges, and for
//! converting formulas to values

use super::{ValueScope, WorkbookModifications, attribute, read_element_body, write_element};
use crate::reader::parser_utils::{attr_value, parse_cell_range, parse_cell_ref, push_general_ref};
use crate::reader::xlsx_parser::{
    is_array_error_placeholder, translate_shared_formula, workbook_part_name,
};
use anyhow::{Context, Result, bail};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::collections::{HashMap, HashSet};
use std::io::{BufRead, BufReader, BufWriter, Cursor, Read, Seek, Write};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

const WORKBOOK_PART: &str = "xl/workbook.xml";
const WORKBOOK_RELS_PART: &str = "xl/_rels/workbook.xml.rels";
const CONTENT_TYPES_PART: &str = "[Content_Types].xml";
/// Part name of the calculation chain, which lists every formula cell.
const CALC_CHAIN_PART: &str = "xl/calcChain.xml";

/// A `<sheet>` entry of `workbook.xml`.
#[derive(Debug, PartialEq)]
struct SheetEntry {
    name: String,
    /// Relationship id of the sheet part in `workbook.xml.rels`.
    rel_id: String,
}

/// Modify an XLSX file by applying specified modifications
pub fn modify_workbook_xlsx<R: Read + Seek, W: Write + Seek>(
    input: R,
    output: W,
    modifications: &WorkbookModifications,
) -> Result<()> {
    let mut archive = ZipArchive::new(input)?;
    let mut zip_writer = ZipWriter::new(output);

    let workbook_xml = read_part(&mut archive, WORKBOOK_PART)?;
    let rels_xml = read_part(&mut archive, WORKBOOK_RELS_PART)?;
    let sheets = parse_sheets(&workbook_xml)?;
    let targets = parse_relationship_targets(&rels_xml)?;
    let sheet_target = |sheet: &SheetEntry| {
        targets.get(&sheet.rel_id).with_context(|| {
            format!(
                "Relationship '{}' not found for sheet '{}'",
                sheet.rel_id, sheet.name
            )
        })
    };

    let removes_sheet = |name: &str| {
        modifications
            .remove_sheets
            .as_ref()
            .is_some_and(|sheets| sheets.contains(name))
    };
    let removed_indices: Vec<usize> = (0..sheets.len())
        .filter(|&i| removes_sheet(&sheets[i].name))
        .collect();
    if !sheets.is_empty() && removed_indices.len() == sheets.len() {
        bail!("Cannot remove every sheet of the workbook");
    }
    let mut removed_targets = HashSet::new();
    let mut removed_parts = HashSet::new();
    for &i in &removed_indices {
        let target = sheet_target(&sheets[i])?;
        let part = workbook_part_name(target);
        removed_parts.insert(sheet_relationships_part(&part));
        removed_parts.insert(part);
        removed_targets.insert(target.as_str());
    }

    // Resolve the worksheet part of every sheet to convert to values
    let mut value_scopes: HashMap<String, Vec<&ValueScope>> = HashMap::new();
    for scope in &modifications.convert_to_values {
        if removes_sheet(&scope.sheet) {
            continue;
        }
        let sheet = sheets
            .iter()
            .find(|sheet| sheet.name == scope.sheet)
            .with_context(|| format!("Sheet '{}' not found in {WORKBOOK_PART}", scope.sheet))?;
        let part = workbook_part_name(sheet_target(sheet)?);
        value_scopes.entry(part).or_default().push(scope);
    }
    // Excel rebuilds the calculation chain when it is missing, but reports a
    // corrupt file when it lists cells that no longer hold formulas or sheets
    // that no longer exist.
    let drop_calc_chain = !value_scopes.is_empty() || !removed_indices.is_empty();
    let removes_part = |part_name: &str| {
        removed_parts.contains(part_name) || (drop_calc_chain && part_name == CALC_CHAIN_PART)
    };
    let removed_ranges = modifications
        .remove_named_ranges
        .as_ref()
        .filter(|ranges| !ranges.is_empty());

    for i in 0..archive.len() {
        let name = archive.by_index_raw(i)?.name().to_string();
        if removes_part(&name) {
            continue;
        }
        let content = match name.as_str() {
            WORKBOOK_PART if !removed_indices.is_empty() || removed_ranges.is_some() => {
                modify_workbook_xml(&workbook_xml, &removed_indices, removed_ranges)?
            }
            WORKBOOK_RELS_PART if drop_calc_chain => remove_relationships(&rels_xml, |target| {
                removed_targets.contains(target) || workbook_part_name(target) == CALC_CHAIN_PART
            })?,
            CONTENT_TYPES_PART if drop_calc_chain => {
                let content = read_part(&mut archive, CONTENT_TYPES_PART)?;
                remove_content_type_overrides(&content, |part_name| {
                    part_name.strip_prefix('/').is_some_and(removes_part)
                })?
            }
            _ => {
                if let Some(scopes) = value_scopes.get(&name) {
                    zip_writer.start_file(&name, SimpleFileOptions::default())?;
                    convert_formulas_to_values(
                        BufReader::new(archive.by_index(i)?),
                        &mut zip_writer,
                        scopes,
                    )
                    .with_context(|| format!("Failed to convert formulas to values in {name}"))?;
                } else {
                    zip_writer.raw_copy_file(archive.by_index_raw(i)?)?;
                }
                continue;
            }
        };
        zip_writer.start_file(&name, SimpleFileOptions::default())?;
        zip_writer.write_all(content.as_bytes())?;
    }

    zip_writer.finish()?;
    Ok(())
}

// Helper functions

fn read_part<R: Read + Seek>(archive: &mut ZipArchive<R>, name: &str) -> Result<String> {
    let mut content = String::new();
    archive
        .by_name(name)
        .with_context(|| format!("Failed to find {name}"))?
        .read_to_string(&mut content)?;
    Ok(content)
}

/// Name of the relationships part of the worksheet part `part`.
fn sheet_relationships_part(part: &str) -> String {
    let (folder, file) = part.rsplit_once('/').unwrap_or(("", part));
    format!("{folder}/_rels/{file}.rels")
}

fn parse_sheets(workbook_xml: &str) -> Result<Vec<SheetEntry>> {
    let mut reader = Reader::from_str(workbook_xml);
    let mut sheets = Vec::new();
    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "sheet" => {
                sheets.push(SheetEntry {
                    name: attribute(&e, "name")?.unwrap_or_default(),
                    rel_id: attribute(&e, "r:id")?.unwrap_or_default(),
                });
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(sheets)
}

/// Map each relationship id of a relationships part to its `Target`.
fn parse_relationship_targets(rels_xml: &str) -> Result<HashMap<String, String>> {
    let mut reader = Reader::from_str(rels_xml);
    let mut targets = HashMap::new();
    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "Relationship" => {
                if let (Some(id), Some(target)) = (attribute(&e, "Id")?, attribute(&e, "Target")?) {
                    targets.insert(id, target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(targets)
}

/// Rewrite `workbook.xml` without the sheets at `removed_indices` (sorted)
/// and the defined names in `removed_ranges`.
///
/// Sheet-local names of the removed sheets go with them, and the sheet
/// indices of the remaining local names and of the workbook views shift past
/// the removed sheets.
fn modify_workbook_xml(
    xml: &str,
    removed_indices: &[usize],
    removed_ranges: Option<&HashSet<String>>,
) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut sheet_index = 0;
    let mut kept_sheets = 0usize;
    // Index of a sheet once the removed ones are gone, if it is kept.
    let new_index = |index: usize| -> Option<usize> {
        match removed_indices.binary_search(&index) {
            Ok(_) => None,
            Err(removed_before) => Some(index - removed_before),
        }
    };

    loop {
        match reader.read_event()? {
            Event::Start(e) if e.name().as_ref() == "sheet" => {
                let end = e.to_end().into_owned();
                if new_index(sheet_index).is_none() {
                    reader.read_to_end(end.name())?;
                } else {
                    kept_sheets += 1;
                    writer.write_event(Event::Start(e))?;
                }
                sheet_index += 1;
            }
            Event::Empty(e) if e.name().as_ref() == "sheet" => {
                if new_index(sheet_index).is_some() {
                    kept_sheets += 1;
                    writer.write_event(Event::Empty(e))?;
                }
                sheet_index += 1;
            }
            Event::Start(e) if e.name().as_ref() == "definedName" => {
                match kept_defined_name(&e, new_index, removed_ranges)? {
                    Some(name) => writer.write_event(Event::Start(name))?,
                    None => {
                        let end = e.to_end().into_owned();
                        reader.read_to_end(end.name())?;
                    }
                }
            }
            Event::Empty(e) if e.name().as_ref() == "definedName" => {
                if let Some(name) = kept_defined_name(&e, new_index, removed_ranges)? {
                    writer.write_event(Event::Empty(name))?;
                }
            }
            Event::Eof => break,
            event => writer.write_event(event)?,
        }
    }
    let xml = String::from_utf8(writer.into_inner().into_inner())?;
    if removed_indices.is_empty() {
        return Ok(xml);
    }
    // Sheets follow the workbook views, so views are rewritten once the
    // number of kept sheets is known.
    remap_workbook_views(&xml, |index| {
        new_index(index)
            .unwrap_or(index - removed_indices.partition_point(|&i| i < index))
            .min(kept_sheets - 1)
    })
}

/// Return the `<definedName>` start tag `element` with its `localSheetId`
/// moved by `new_index`, or `None` when the name is removed or local to a
/// removed sheet.
fn kept_defined_name<'a>(
    element: &BytesStart<'a>,
    new_index: impl Fn(usize) -> Option<usize>,
    removed_ranges: Option<&HashSet<String>>,
) -> Result<Option<BytesStart<'a>>> {
    let name = attribute(element, "name")?.unwrap_or_default();
    if removed_ranges.is_some_and(|ranges| ranges.contains(&name)) {
        return Ok(None);
    }
    let Some(local_index) = attribute(element, "localSheetId")? else {
        return Ok(Some(element.clone()));
    };
    new_index(local_index.parse()?)
        .map(|index| with_attribute(element, "localSheetId", &index.to_string()))
        .transpose()
}

/// Rewrite the `activeTab` and `firstSheet` sheet indices of the workbook
/// views of `xml` with `remap`.
fn remap_workbook_views(xml: &str, remap: impl Fn(usize) -> usize) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    loop {
        match reader.read_event()? {
            Event::Empty(e) if e.name().as_ref() == "workbookView" => {
                let mut view = e.clone();
                for key in ["firstSheet", "activeTab"] {
                    if let Some(index) = attribute(&e, key)? {
                        view = with_attribute(&view, key, &remap(index.parse()?).to_string())?;
                    }
                }
                writer.write_event(Event::Empty(view))?;
            }
            Event::Eof => break,
            event => writer.write_event(event)?,
        }
    }
    Ok(String::from_utf8(writer.into_inner().into_inner())?)
}

/// Copy `element` with the attribute `key` set to `value`, keeping the
/// attribute order.
fn with_attribute<'a>(element: &BytesStart<'a>, key: &str, value: &str) -> Result<BytesStart<'a>> {
    let mut rebuilt = element.clone();
    rebuilt.clear_attributes();
    for attr in element.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == key {
            rebuilt.push_attribute((key, value));
        } else {
            rebuilt.push_attribute(attr);
        }
    }
    Ok(rebuilt)
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

/// Formula text of a shared group whose first cell is converted to its
/// value, which leaves the rest of the group without formula text.
struct ConvertedSharedFormula {
    row: u32,
    col: u32,
    text: String,
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

/// Read the formula of a cell and whether it holds a cached value, as the
/// reader does: an array error placeholder is no value.
fn inspect_cell(start: &BytesStart, body: &[Event]) -> Result<(Option<CellFormula>, bool)> {
    let mut formula: Option<CellFormula> = None;
    let mut value: Option<String> = None;
    let mut in_formula = false;
    let mut in_value = false;
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
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "v" => {
                value = Some(String::new());
                in_value = matches!(event, Event::Start(_));
            }
            Event::End(e) if e.name().as_ref() == "v" => in_value = false,
            Event::Text(text) => {
                if let Some(text_buf) = text_target(&mut formula, &mut value, in_formula, in_value)
                {
                    text_buf.push_str(text);
                }
            }
            Event::CData(text) => {
                if let Some(text_buf) = text_target(&mut formula, &mut value, in_formula, in_value)
                {
                    text_buf.push_str(text);
                }
            }
            Event::GeneralRef(reference) => {
                if let Some(text_buf) = text_target(&mut formula, &mut value, in_formula, in_value)
                {
                    push_general_ref(text_buf, reference)?;
                }
            }
            _ => {}
        }
    }
    let placeholder = attribute(start, "t")?.as_deref() == Some("e")
        && value
            .as_deref()
            .zip(formula.as_ref())
            .is_some_and(|(value, formula)| {
                !formula.text.is_empty()
                    && is_array_error_placeholder(
                        value.trim(),
                        &formula.text,
                        formula.kind.as_deref() == Some("array"),
                    )
            });
    Ok((formula, value.is_some() && !placeholder))
}

/// The text being collected: the formula's or the cached value's.
fn text_target<'a>(
    formula: &'a mut Option<CellFormula>,
    value: &'a mut Option<String>,
    in_formula: bool,
    in_value: bool,
) -> Option<&'a mut String> {
    if in_formula {
        formula.as_mut().map(|formula| &mut formula.text)
    } else if in_value {
        value.as_mut()
    } else {
        None
    }
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
            scopes
                .iter()
                .any(|scope| scope.contains_range((first_row, first_col), (last_row, last_col)))
        }
        _ => scopes.iter().any(|scope| scope.contains(row, col)),
    }
}

/// Replace the formulas of the cells in `scopes` by their cached values.
///
/// A cell of a shared group whose first cell is converted, but that is not
/// converted itself, gets its own copy of the translated formula. The first
/// cell of a group precedes the rest in the worksheet.
fn convert_formulas_to_values<R: BufRead, W: Write>(
    input: R,
    output: W,
    scopes: &[&ValueScope],
) -> Result<()> {
    let mut reader = Reader::from_reader(input);
    let mut writer = Writer::new(BufWriter::new(output));
    let mut cursor = CellCursor::default();
    let mut converted_shared: HashMap<u32, ConvertedSharedFormula> = HashMap::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf)? {
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
                let body = read_element_body(&mut reader, "c")?;
                let (formula, has_value) = inspect_cell(&e, &body)?;
                let Some(formula) = formula else {
                    write_element(&mut writer, e, Some(&body))?;
                    continue;
                };
                let converts = converts_to_value(scopes, row, col, &formula, has_value);
                if let Some(si) = formula.shared_index.filter(|_| formula.is_shared()) {
                    if !formula.text.is_empty() {
                        if converts {
                            converted_shared.insert(
                                si,
                                ConvertedSharedFormula {
                                    row,
                                    col,
                                    text: formula.text,
                                },
                            );
                        } else {
                            converted_shared.remove(&si);
                        }
                    } else if !converts && let Some(group) = converted_shared.get(&si) {
                        let text = translate_shared_formula(
                            &group.text,
                            row.checked_signed_diff(group.row)
                                .context("shared formula row out of range")?,
                            col.checked_signed_diff(group.col)
                                .context("shared formula column out of range")?,
                        );
                        write_unshared_cell(&mut writer, e, &body, &text)?;
                        continue;
                    }
                }
                if converts {
                    write_value_cell(&mut writer, &e, &body)?;
                } else {
                    write_element(&mut writer, e, Some(&body))?;
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

    writer.into_inner().flush()?;
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

    const WORKBOOK_XML: &str = r#"<workbook><sheets><sheet name="R&amp;D" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="A&amp;B">Data!$A$1</definedName><definedName name="Keep">Data!$B$1</definedName></definedNames></workbook>"#;

    #[test]
    fn test_parse_sheets_unescapes_names() {
        let sheets = parse_sheets(WORKBOOK_XML).unwrap();
        assert_eq!(
            sheets,
            vec![
                SheetEntry {
                    name: "R&D".to_string(),
                    rel_id: "rId1".to_string()
                },
                SheetEntry {
                    name: "Data".to_string(),
                    rel_id: "rId2".to_string()
                }
            ]
        );
    }

    #[test]
    fn test_remove_sheet_with_escaped_name() {
        let xml = modify_workbook_xml(WORKBOOK_XML, &[0], None).unwrap();
        assert!(!xml.contains("R&amp;D"));
        assert!(xml.contains(r#"<sheet name="Data" sheetId="2" r:id="rId2"/>"#));
    }

    #[test]
    fn test_remove_named_range_with_escaped_name() {
        let remove = HashSet::from(["A&B".to_string()]);
        let xml = modify_workbook_xml(WORKBOOK_XML, &[], Some(&remove)).unwrap();
        assert!(!xml.contains("A&amp;B"));
        assert!(xml.contains(r#"<definedName name="Keep">Data!$B$1</definedName>"#));
    }

    #[test]
    fn test_remove_empty_defined_names() {
        let xml = r#"<workbook><sheets><sheet name="A"/><sheet name="B"/></sheets><definedNames><definedName name="Drop"/><definedName name="Local" localSheetId="0"/><definedName name="Moved" localSheetId="1"/></definedNames></workbook>"#;
        let remove = HashSet::from(["Drop".to_string()]);
        let modified = modify_workbook_xml(xml, &[0], Some(&remove)).unwrap();
        assert!(
            modified.contains(
                r#"<definedNames><definedName name="Moved" localSheetId="0"/></definedNames>"#
            ),
            "{modified}"
        );
    }

    #[test]
    fn test_workbook_views_move_to_a_kept_sheet() {
        let xml = r#"<workbook><bookViews><workbookView firstSheet="1" activeTab="3"/></bookViews><sheets><sheet name="A"/><sheet name="B"/><sheet name="C"/><sheet name="D"/></sheets></workbook>"#;
        let modified = modify_workbook_xml(xml, &[1, 3], None).unwrap();
        assert!(
            modified.contains(r#"<workbookView firstSheet="1" activeTab="1"/>"#),
            "{modified}"
        );
    }

    #[test]
    fn test_sheet_relationships_part() {
        assert_eq!(
            sheet_relationships_part("xl/worksheets/sheet2.xml"),
            "xl/worksheets/_rels/sheet2.xml.rels"
        );
    }
}

//! XLSX writer functionality for removing sheets and named ranges, and for
//! converting formulas to values

use super::structure::StructureEdits;
use super::xlsx_structure::{
    Handled, SheetStructure, handle_element, map_attributes, rewrite_comments, rewrite_drawing,
    rewrite_vml, text_of,
};
use super::{ValueScope, WorkbookModifications, attribute, read_element_body, write_element};
use crate::reader::parser_utils::{
    attr_value, column_letters, parse_cell_range, parse_cell_ref, push_general_ref,
};
use crate::reader::xlsx_parser::{
    is_array_error_placeholder, translate_shared_formula, workbook_part_name,
};
use anyhow::{Context, Result, bail};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use sheetrs_formula::edit::Edit;
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
    let sheet_names: Vec<String> = sheets.iter().map(|sheet| sheet.name.clone()).collect();
    let removed_sheet_names: HashSet<String> = removed_indices
        .iter()
        .map(|&i| sheets[i].name.clone())
        .collect();
    let plan = StructureEdits::new(
        Some(&sheet_names),
        &removed_sheet_names,
        &modifications.delete_rows,
        &modifications.delete_columns,
    )?;
    let edits = plan.edits();
    // The worksheet parts of the kept sheets, by sheet name.
    let mut sheet_of_part: HashMap<String, &str> = HashMap::new();
    for (i, sheet) in sheets.iter().enumerate() {
        if !removed_indices.contains(&i) {
            sheet_of_part.insert(workbook_part_name(sheet_target(sheet)?), &sheet.name);
        }
    }
    if !plan.is_empty() {
        refuse_what_is_not_adjusted(&mut archive)?;
    }
    // The parts that hold the notes of a sheet whose cells move.
    let mut comment_parts: HashMap<String, (CommentPart, &str)> = HashMap::new();
    for (part, &sheet) in &sheet_of_part {
        if !plan.rows_of(sheet).is_empty() || !plan.cols_of(sheet).is_empty() {
            let rels_part = sheet_relationships_part(part);
            if archive.index_for_name(&rels_part).is_none() {
                continue;
            }
            let rels = read_part(&mut archive, &rels_part)?;
            let base = part.rsplit_once('/').map_or("", |(folder, _)| folder);
            for (kind, target) in parse_typed_relationships(&rels)? {
                let comment_part = match kind.rsplit('/').next().unwrap_or("") {
                    "comments" => CommentPart::Comments,
                    "threadedComment" => CommentPart::Threaded,
                    "vmlDrawing" => CommentPart::Vml,
                    "drawing" => CommentPart::Drawing,
                    _ => continue,
                };
                comment_parts.insert(resolve_part(base, &target), (comment_part, sheet));
            }
        }
    }
    // Excel rebuilds the calculation chain when it is missing, but reports a
    // corrupt file when it lists cells that no longer hold formulas or sheets
    // that no longer exist.
    let drop_calc_chain =
        !value_scopes.is_empty() || !removed_indices.is_empty() || !plan.is_empty();
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
            WORKBOOK_PART
                if !removed_indices.is_empty()
                    || removed_ranges.is_some()
                    || !plan.is_empty()
                    || modifications.recalculate_on_open =>
            {
                modify_workbook_xml(
                    &workbook_xml,
                    &WorkbookEdits {
                        removed_indices: &removed_indices,
                        removed_ranges,
                        plan: &plan,
                        edits: &edits,
                        sheet_names: &sheet_names,
                        recalculate_on_open: modifications.recalculate_on_open,
                    },
                )?
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
                let scopes = value_scopes
                    .get(&name)
                    .map(Vec::as_slice)
                    .unwrap_or_default();
                let structure =
                    sheet_of_part
                        .get(&name)
                        .filter(|_| !plan.is_empty())
                        .map(|&sheet| SheetStructure {
                            plan: &plan,
                            edits: &edits,
                            sheet: sheet.to_string(),
                            rows: plan.rows_of(sheet),
                            cols: plan.cols_of(sheet),
                            recalculate: false,
                        });
                // A sheet that names none of the sheets edited holds no formula that changes.
                let unaffected = match &structure {
                    Some(s) if scopes.is_empty() && !s.changes_layout() => plan
                        .text_names_no_edited_sheet(BufReader::new(archive.by_index(i)?))?
                        .unwrap_or(false),
                    _ => false,
                };
                if unaffected {
                    zip_writer.raw_copy_file(archive.by_index_raw(i)?)?;
                } else if !scopes.is_empty() || structure.is_some() {
                    zip_writer.start_file(&name, SimpleFileOptions::default())?;
                    transform_sheet(
                        BufReader::new(archive.by_index(i)?),
                        &mut zip_writer,
                        scopes,
                        structure.as_ref(),
                        sheet_of_part.get(&name).copied().unwrap_or_default(),
                    )
                    .with_context(|| format!("Failed to rewrite {name}"))?;
                } else if let Some(&(kind, sheet)) = comment_parts.get(&name) {
                    let content = read_part(&mut archive, &name)?;
                    let structure = SheetStructure {
                        plan: &plan,
                        edits: &edits,
                        sheet: sheet.to_string(),
                        rows: plan.rows_of(sheet),
                        cols: plan.cols_of(sheet),
                        recalculate: false,
                    };
                    let new = match kind {
                        CommentPart::Comments => rewrite_comments(&content, "comment", &structure)?,
                        CommentPart::Threaded => {
                            rewrite_comments(&content, "threadedComment", &structure)?
                        }
                        CommentPart::Vml => rewrite_vml(&content, &structure)?,
                        CommentPart::Drawing => rewrite_drawing(&content, &structure)?,
                    };
                    zip_writer.start_file(&name, SimpleFileOptions::default())?;
                    zip_writer.write_all(new.as_bytes())?;
                } else if !plan.is_empty() && name.starts_with("xl/charts/chart") {
                    let content = read_part(&mut archive, &name)?;
                    zip_writer.start_file(&name, SimpleFileOptions::default())?;
                    zip_writer
                        .write_all(rewrite_chart_references(&content, &plan, &edits)?.as_bytes())?;
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

/// A part that holds the notes of a sheet.
#[derive(Clone, Copy)]
enum CommentPart {
    Comments,
    Threaded,
    Vml,
    /// The shapes, images and charts of the sheet, which are anchored to cells.
    Drawing,
}

/// The `(Type, Target)` of each relationship of a relationships part.
fn parse_typed_relationships(rels_xml: &str) -> Result<Vec<(String, String)>> {
    let mut reader = Reader::from_str(rels_xml);
    let mut out = Vec::new();
    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "Relationship" => {
                if let (Some(kind), Some(target)) =
                    (attribute(&e, "Type")?, attribute(&e, "Target")?)
                {
                    out.push((kind, target));
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(out)
}

/// The part name that `target`, as a relationship of a part in the folder `base`, points to.
fn resolve_part(base: &str, target: &str) -> String {
    if let Some(absolute) = target.strip_prefix('/') {
        return absolute.to_string();
    }
    let mut parts: Vec<&str> = base.split('/').filter(|p| !p.is_empty()).collect();
    for segment in target.split('/') {
        match segment {
            ".." => {
                parts.pop();
            }
            "." | "" => {}
            other => parts.push(other),
        }
    }
    parts.join("/")
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

/// What a command changes in `workbook.xml`.
struct WorkbookEdits<'a> {
    /// Indices of the removed sheets, sorted.
    removed_indices: &'a [usize],
    /// Names of the removed named ranges.
    removed_ranges: Option<&'a HashSet<String>>,
    plan: &'a StructureEdits,
    edits: &'a [Edit<'a>],
    /// Names of the sheets of the input, in order.
    sheet_names: &'a [String],
    recalculate_on_open: bool,
}

/// Rewrite `workbook.xml` without the removed sheets and named ranges, with the references of the
/// names that stay adjusted to the deleted rows, columns and sheets.
///
/// Sheet-local names of the removed sheets go with them, and the sheet
/// indices of the remaining local names and of the workbook views shift past
/// the removed sheets. With `recalculate_on_open` the application that opens the
/// workbook recalculates it.
fn modify_workbook_xml(xml: &str, w: &WorkbookEdits) -> Result<String> {
    let removed_indices = w.removed_indices;
    let removed_ranges = w.removed_ranges;
    let rewrite_names = !w.plan.is_empty() || removed_ranges.is_some();
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut sheet_index = 0;
    let mut kept_sheets = 0usize;
    let mut calc_pr = false;
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
                    Some(name) if rewrite_names => {
                        let home = attribute(&e, "localSheetId")?
                            .and_then(|i| i.parse::<usize>().ok())
                            .and_then(|i| w.sheet_names.get(i))
                            .map_or("", String::as_str);
                        let label = attribute(&e, "name")?.unwrap_or_default();
                        let end = e.to_end().into_owned();
                        let body = read_element_body(&mut reader, "definedName")?;
                        let text = text_of(&body)?;
                        let new = definition_after(w, &text, home)
                            .with_context(|| format!("Failed to adjust the name '{label}'"))?;
                        writer.write_event(Event::Start(name))?;
                        writer.write_event(Event::Text(BytesText::new(&new)))?;
                        writer.write_event(Event::End(end))?;
                    }
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
            Event::Empty(e) if e.name().as_ref() == "calcPr" && w.recalculate_on_open => {
                calc_pr = true;
                writer.write_event(Event::Empty(full_calc_on_load(&e)?))?;
            }
            Event::Start(e) if e.name().as_ref() == "calcPr" && w.recalculate_on_open => {
                calc_pr = true;
                writer.write_event(Event::Start(full_calc_on_load(&e)?))?;
            }
            // `calcPr` goes before these elements of the schema, or at the end.
            Event::Start(e)
                if w.recalculate_on_open && !calc_pr && follows_calc_pr(e.name().as_ref()) =>
            {
                calc_pr = true;
                writer.write_event(Event::Empty(full_calc_on_load_new()))?;
                writer.write_event(Event::Start(e))?;
            }
            Event::Empty(e)
                if w.recalculate_on_open && !calc_pr && follows_calc_pr(e.name().as_ref()) =>
            {
                calc_pr = true;
                writer.write_event(Event::Empty(full_calc_on_load_new()))?;
                writer.write_event(Event::Empty(e))?;
            }
            Event::End(e) if e.name().as_ref() == "workbook" => {
                if w.recalculate_on_open && !calc_pr {
                    writer.write_event(Event::Empty(full_calc_on_load_new()))?;
                }
                writer.write_event(Event::End(e))?;
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

/// A `calcPr` that asks for a recalculation when the workbook is opened.
fn full_calc_on_load_new() -> BytesStart<'static> {
    let mut element = BytesStart::new("calcPr");
    element.push_attribute(("fullCalcOnLoad", "1"));
    element
}

/// Whether the workbook element `name` comes after `calcPr` in the schema.
fn follows_calc_pr(name: &str) -> bool {
    matches!(
        name,
        "oleSize"
            | "customWorkbookViews"
            | "pivotCaches"
            | "smartTagPr"
            | "smartTagTypes"
            | "webPublishing"
            | "fileRecoveryPr"
            | "webPublishObjects"
            | "extLst"
    )
}

/// `calcPr` with `fullCalcOnLoad` on.
fn full_calc_on_load(element: &BytesStart) -> Result<BytesStart<'static>> {
    let mut seen = false;
    let mut rebuilt = map_attributes(element, |key, _| {
        Ok((key == "fullCalcOnLoad").then(|| {
            seen = true;
            Some("1".to_string())
        }))
    })?;
    if !seen {
        rebuilt.push_attribute(("fullCalcOnLoad", "1"));
    }
    Ok(rebuilt)
}

/// The definition of a name once the plan is applied: its references adjusted, and the removed
/// names it uses unknown (`#NAME?`), as LibreOffice leaves them.
fn definition_after(w: &WorkbookEdits, text: &str, home: &str) -> Result<String> {
    let adjusted = w
        .plan
        .rewrite_formula(w.edits, text, home)?
        .unwrap_or_else(|| text.to_string());
    Ok(match w.removed_ranges {
        Some(removed) => replace_names(&adjusted, removed),
        None => adjusted,
    })
}

/// `text` with the uses of the names in `names` (without regard to case) replaced by `#NAME?`.
/// Text between quotes, calls and sheet qualifiers are left alone.
pub(super) fn replace_names(text: &str, names: &HashSet<String>) -> String {
    let lower: HashSet<String> = names.iter().map(|name| name.to_lowercase()).collect();
    let bytes = text.as_bytes();
    let mut out = String::with_capacity(text.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'"' {
            let mut j = i + 1;
            while let Some(quote) = text[j..].find('"') {
                j += quote + 1;
                if bytes.get(j) != Some(&b'"') {
                    break;
                }
                j += 1;
            }
            let end = if text[i + 1..].contains('"') {
                j
            } else {
                text.len()
            };
            out.push_str(&text[i..end]);
            i = end;
        } else if bytes[i] == b'\'' {
            // A quoted sheet name: its words are not names.
            let mut j = i + 1;
            while let Some(quote) = text[j..].find('\'') {
                j += quote + 1;
                if bytes.get(j) != Some(&b'\'') {
                    break;
                }
                j += 1;
            }
            out.push_str(&text[i..j]);
            i = j;
        } else if bytes[i].is_ascii_alphabetic()
            || matches!(bytes[i], b'_' | b'\\')
            || !bytes[i].is_ascii()
        {
            let mut j = i + 1;
            while j < bytes.len()
                && (bytes[j].is_ascii_alphanumeric()
                    || matches!(bytes[j], b'_' | b'.')
                    || !bytes[j].is_ascii())
            {
                j += 1;
            }
            let token = &text[i..j];
            let used_as_name =
                !matches!(bytes.get(j), Some(b'(' | b'!')) && lower.contains(&token.to_lowercase());
            out.push_str(if used_as_name { "#NAME?" } else { token });
            i = j;
        } else {
            let ch = text[i..].chars().next().expect("a character at a boundary");
            out.push(ch);
            i += ch.len_utf8();
        }
    }
    out
}

/// Refuse the parts that hold references the command would leave out of date.
fn refuse_what_is_not_adjusted<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<()> {
    for i in 0..archive.len() {
        let name = archive.by_index_raw(i)?.name().to_string();
        if name.starts_with("xl/pivotTables/") || name.starts_with("xl/pivotCache/") {
            bail!("the workbook has pivot tables, whose sources are not adjusted yet");
        }
    }
    Ok(())
}

/// A chart part with the ranges of its series adjusted.
fn rewrite_chart_references(
    xml: &str,
    plan: &StructureEdits,
    edits: &[Edit<'_>],
) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    loop {
        match reader.read_event()? {
            Event::Start(e) if e.name().as_ref().ends_with(":f") => {
                let name = e.name().as_ref().to_string();
                let body = read_element_body(&mut reader, &name)?;
                let text = text_of(&body)?;
                // A reference that goes becomes `#REF!`, as a chart of LibreOffice shows it.
                let new = plan.rewrite_formula(edits, &text, "")?.unwrap_or(text);
                let end = e.to_end().into_owned();
                writer.write_event(Event::Start(e))?;
                writer.write_event(Event::Text(BytesText::new(&new)))?;
                writer.write_event(Event::End(end))?;
            }
            Event::Eof => break,
            event => writer.write_event(event)?,
        }
    }
    Ok(String::from_utf8(writer.into_inner().into_inner())?)
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

/// The first cell of a shared group, which the rest of the group takes its formula from.
struct SharedMaster {
    row: u32,
    col: u32,
    text: String,
    /// The master was replaced by its value: the group has no formula text left.
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

/// How a formula cell writes its `<f>` element.
enum FormulaOut {
    /// As it was read.
    Keep,
    /// A plain formula of its own, which a cell of a shared group gets when its text changes.
    Plain(String),
    /// An array formula over `range`.
    Array { text: String, range: String },
}

/// Rewrite a worksheet: replace the formulas of the cells in `scopes` by their cached values, and
/// apply `structure` (rows and columns deleted, references to them and to removed sheets
/// adjusted). Both read the coordinates of the input.
///
/// A cell of a shared group whose first cell is converted, or whose formula text changes, gets its
/// own copy of the translated formula. The first cell of a group precedes the rest in the
/// worksheet. A formula that changes loses its cached value, which is no longer the result.
fn transform_sheet<R: BufRead, W: Write>(
    input: R,
    output: W,
    scopes: &[&ValueScope],
    structure: Option<&SheetStructure>,
    sheet: &str,
) -> Result<()> {
    let mut reader = Reader::from_reader(input);
    let mut writer = Writer::new(BufWriter::new(output));
    let mut cursor = CellCursor::default();
    let mut masters: HashMap<u32, SharedMaster> = HashMap::new();
    let layout = structure.filter(|s| s.changes_layout());
    // Inside `rowBreaks` (true) or `colBreaks` (false).
    let mut break_rows: Option<bool> = None;
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name().as_ref() == "row" => {
                cursor.start_row(&e)?;
                match layout.map(|s| s.rows.moved(cursor.row)) {
                    Some(None) => {
                        skip_row(&mut reader, structure, &e, cursor.row, &mut masters)?;
                    }
                    Some(Some(new_row)) => {
                        writer.write_event(Event::Start(renumbered_row(&e, new_row)?))?;
                    }
                    None => writer.write_event(Event::Start(e))?,
                }
            }
            Event::Empty(e) if e.name().as_ref() == "row" => {
                cursor.start_row(&e)?;
                match layout.map(|s| s.rows.moved(cursor.row)) {
                    Some(None) => {}
                    Some(Some(new_row)) => {
                        writer.write_event(Event::Empty(renumbered_row(&e, new_row)?))?;
                    }
                    None => writer.write_event(Event::Empty(e))?,
                }
            }
            Event::Start(e) if e.name().as_ref() == "c" => {
                let (row, col) = cursor.next_cell(&e)?;
                let body = read_element_body(&mut reader, "c")?;
                let (formula, has_value) = inspect_cell(&e, &body)?;
                // The cells of a group read their formula from the first one, even when it goes.
                if let Some(f) = &formula
                    && f.is_shared()
                    && let Some(si) = f.shared_index
                    && !f.text.is_empty()
                {
                    masters.insert(
                        si,
                        SharedMaster {
                            row,
                            col,
                            text: f.text.clone(),
                            converted: false,
                        },
                    );
                }
                let at = match layout {
                    Some(s) => match s.cols.moved(col) {
                        Some(new_col) => Some((
                            s.rows
                                .moved(row)
                                .expect("a cell of a deleted row is skipped with its row"),
                            new_col,
                        )),
                        None => {
                            check_array_goes_whole(structure, formula.as_ref(), row, col)?;
                            continue;
                        }
                    },
                    None => None,
                };
                let start = relocated_cell(&e, at, false)?;
                let Some(formula) = formula else {
                    write_element(&mut writer, start, Some(&body))?;
                    continue;
                };
                let converts = converts_to_value(scopes, row, col, &formula, has_value);
                let shared_index = formula.shared_index.filter(|_| formula.is_shared());
                if let Some(si) = shared_index
                    && !formula.text.is_empty()
                {
                    masters.insert(
                        si,
                        SharedMaster {
                            row,
                            col,
                            text: formula.text.clone(),
                            converted: converts,
                        },
                    );
                }
                if converts {
                    write_value_cell(&mut writer, &start, &body)?;
                    continue;
                }
                let (out, changed) = match structure {
                    Some(s) => rewritten_formula(s, &formula, row, col, &masters)?,
                    None => (FormulaOut::Keep, false),
                };
                // A cell of a group whose first cell became a value needs a formula of its own.
                let out = match (out, shared_index) {
                    (FormulaOut::Keep, Some(si)) if formula.text.is_empty() => {
                        match masters.get(&si).filter(|m| m.converted) {
                            Some(m) => FormulaOut::Plain(shared_text(m, row, col, sheet)?),
                            None => FormulaOut::Keep,
                        }
                    }
                    (out, _) => out,
                };
                let start = relocated_cell(&e, at, changed)?;
                match out {
                    FormulaOut::Keep => write_element(&mut writer, start, Some(&body))?,
                    out => write_formula_cell(&mut writer, start, &body, &out, changed)?,
                }
            }
            Event::Empty(e) if e.name().as_ref() == "c" => {
                let (_, col) = cursor.next_cell(&e)?;
                match layout {
                    Some(s) => {
                        if let Some(new_col) = s.cols.moved(col) {
                            let row = s
                                .rows
                                .moved(cursor.row)
                                .expect("a cell of a deleted row is skipped with its row");
                            writer.write_event(Event::Empty(relocated_cell(
                                &e,
                                Some((row, new_col)),
                                false,
                            )?))?;
                        }
                    }
                    None => writer.write_event(Event::Empty(e))?,
                }
            }
            Event::Start(e)
                if layout.is_some() && matches!(e.name().as_ref(), "rowBreaks" | "colBreaks") =>
            {
                // The counts are optional, and would be out of date.
                break_rows = Some(e.name().as_ref() == "rowBreaks");
                let breaks = map_attributes(&e, |key, _| {
                    Ok(matches!(key, "count" | "manualBreakCount").then_some(None))
                })?;
                writer.write_event(Event::Start(breaks))?;
            }
            Event::Empty(e)
                if e.name().as_ref() == "brk" && layout.is_some() && break_rows.is_some() =>
            {
                // `id` is the 0-based row (or column) after the break: the break goes with it.
                let s = layout.expect("matched above");
                let bands = if break_rows == Some(true) {
                    &s.rows
                } else {
                    &s.cols
                };
                let id = attribute(&e, "id")?
                    .unwrap_or_default()
                    .parse::<u32>()
                    .unwrap_or(0);
                if let Some(moved) = bands.moved(id) {
                    let moved_break = map_attributes(&e, |key, _| {
                        Ok((key == "id").then(|| Some(moved.to_string())))
                    })?;
                    writer.write_event(Event::Empty(moved_break))?;
                }
            }
            Event::End(e) if matches!(e.name().as_ref(), "rowBreaks" | "colBreaks") => {
                break_rows = None;
                writer.write_event(Event::End(e))?;
            }
            Event::Start(e) if structure.is_some() => {
                let s = structure.expect("matched above");
                if let Handled::No = handle_element(&mut reader, &mut writer, &e, true, s)? {
                    writer.write_event(Event::Start(e))?;
                }
            }
            Event::Empty(e) if structure.is_some() => {
                let s = structure.expect("matched above");
                if let Handled::No = handle_element(&mut reader, &mut writer, &e, false, s)? {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Event::Eof => break,
            event => writer.write_event(event)?,
        }
    }

    writer.into_inner().flush()?;
    Ok(())
}

/// The text of the formula of the cell `(row, col)` of the shared group of `master`.
fn shared_text(master: &SharedMaster, row: u32, col: u32, sheet: &str) -> Result<String> {
    Ok(translate_shared_formula(
        &master.text,
        row.checked_signed_diff(master.row)
            .context("shared formula row out of range")?,
        col.checked_signed_diff(master.col)
            .context("shared formula column out of range")?,
        sheet,
    ))
}

/// What a formula cell writes once the plan is applied, and whether its text changed.
fn rewritten_formula(
    s: &SheetStructure,
    formula: &CellFormula,
    row: u32,
    col: u32,
    masters: &HashMap<u32, SharedMaster>,
) -> Result<(FormulaOut, bool)> {
    let shared = formula.shared_index.filter(|_| formula.is_shared());
    let original = match shared {
        Some(_) if !formula.text.is_empty() => formula.text.clone(),
        Some(si) => match masters.get(&si) {
            Some(master) => shared_text(master, row, col, &s.sheet)?,
            None => bail!(
                "the cell {}{} of sheet '{}' shares a formula whose first cell is not in the file",
                column_letters(col),
                row + 1,
                s.sheet
            ),
        },
        None => formula.text.clone(),
    };
    if formula.kind.as_deref() == Some("dataTable") && s.changes_layout() {
        bail!(
            "sheet '{}' has a data table, whose cells are not adjusted yet",
            s.sheet
        );
    }
    let rewritten = s
        .rewrite(&original)
        .with_context(|| format!("formula in {}!{}{}", s.sheet, column_letters(col), row + 1))?;
    let changed = rewritten.is_some();
    let text = rewritten.unwrap_or(original);
    if formula.kind.as_deref() == Some("array") {
        let (r0, c0, r1, c1) = formula.range.unwrap_or((row, col, row, col));
        let range = if s.changes_layout() {
            let new = s.plan.range(&s.sheet, r0, c0, r1, c1);
            // Every cell of an array goes or stays: a part of it cannot be deleted.
            match new {
                Some((nr0, nc0, nr1, nc1)) if nr1 - nr0 == r1 - r0 && nc1 - nc0 == c1 - c0 => {
                    format!(
                        "{}{}:{}{}",
                        column_letters(nc0),
                        nr0 + 1,
                        column_letters(nc1),
                        nr1 + 1
                    )
                }
                _ => bail!(
                    "a deletion in sheet '{}' covers part of the array formula at {}{}",
                    s.sheet,
                    column_letters(c0),
                    r0 + 1
                ),
            }
        } else {
            format!(
                "{}{}:{}{}",
                column_letters(c0),
                r0 + 1,
                column_letters(c1),
                r1 + 1
            )
        };
        return Ok((FormulaOut::Array { text, range }, changed));
    }
    let expand = shared.is_some() && (s.changes_layout() || changed);
    Ok(match (changed, expand) {
        (false, false) => (FormulaOut::Keep, false),
        _ => (FormulaOut::Plain(text), changed),
    })
}

/// An array formula whose first cell goes cannot keep the rest of its cells.
fn check_array_goes_whole(
    structure: Option<&SheetStructure>,
    formula: Option<&CellFormula>,
    row: u32,
    col: u32,
) -> Result<()> {
    let (Some(s), Some(formula)) = (structure, formula) else {
        return Ok(());
    };
    if formula.kind.as_deref() == Some("array") {
        let (r0, c0, r1, c1) = formula.range.unwrap_or((row, col, row, col));
        if s.plan.range(&s.sheet, r0, c0, r1, c1).is_some() {
            bail!(
                "a deletion in sheet '{}' covers part of the array formula at {}{}",
                s.sheet,
                column_letters(c0),
                r0 + 1
            );
        }
    }
    Ok(())
}

/// Skip the row `row` that starts at `start`, which the plan deletes. The first cell of a shared
/// group in it stays known for the cells of the group in other rows, and an array formula that
/// starts in it must go whole.
fn skip_row<R: BufRead>(
    reader: &mut Reader<R>,
    structure: Option<&SheetStructure>,
    start: &BytesStart,
    row: u32,
    masters: &mut HashMap<u32, SharedMaster>,
) -> Result<()> {
    let end = start.to_end().into_owned();
    let body = read_element_body(reader, end.name().as_ref())?;
    let mut i = 0;
    while i < body.len() {
        let Event::Start(cell) = &body[i] else {
            i += 1;
            continue;
        };
        if cell.name().as_ref() != "c" {
            i += 1;
            continue;
        }
        let mut stop = i + 1;
        let mut depth = 0u32;
        while stop < body.len() {
            match &body[stop] {
                Event::Start(e) if e.name().as_ref() == "c" => depth += 1,
                Event::End(e) if e.name().as_ref() == "c" => {
                    if depth == 0 {
                        break;
                    }
                    depth -= 1;
                }
                _ => {}
            }
            stop += 1;
        }
        let col = attribute(cell, "r")?
            .and_then(|r| parse_cell_ref(&r))
            .map_or(0, |(_, col)| col);
        let (formula, _) = inspect_cell(cell, &body[i + 1..stop])?;
        if let Some(formula) = formula {
            if formula.is_shared()
                && let Some(si) = formula.shared_index
                && !formula.text.is_empty()
            {
                masters.insert(
                    si,
                    SharedMaster {
                        row,
                        col,
                        text: formula.text.clone(),
                        converted: false,
                    },
                );
            }
            check_array_goes_whole(structure, Some(&formula), row, col)?;
        }
        i = stop + 1;
    }
    Ok(())
}

/// A row element with its new number; its `spans` hint no longer holds.
fn renumbered_row(row: &BytesStart, new_row: u32) -> Result<BytesStart<'static>> {
    map_attributes(row, |key, _| {
        Ok(match key {
            "r" => Some(Some((new_row + 1).to_string())),
            "spans" => Some(None),
            _ => None,
        })
    })
}

/// A cell element at `at` (`(row, col)`, 0-based) when it moves, without the type of a result that
/// is no longer there when `drop_type`.
fn relocated_cell(
    cell: &BytesStart,
    at: Option<(u32, u32)>,
    drop_type: bool,
) -> Result<BytesStart<'static>> {
    map_attributes(cell, |key, _| {
        Ok(match (key, at) {
            ("r", Some((row, col))) => Some(Some(format!("{}{}", column_letters(col), row + 1))),
            ("t", _) if drop_type => Some(None),
            _ => None,
        })
    })
}

/// Write a formula cell with the formula `out`, and without its cached value when `drop_value`.
fn write_formula_cell<W: Write>(
    writer: &mut Writer<W>,
    start: BytesStart,
    body: &[Event],
    out: &FormulaOut,
    drop_value: bool,
) -> Result<()> {
    let end = start.to_end().into_owned();
    writer.write_event(Event::Start(start))?;
    let (mut in_formula, mut in_value) = (false, false);
    for event in body {
        match event {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "f" => {
                let (text, range) = match out {
                    FormulaOut::Plain(text) => (text, None),
                    FormulaOut::Array { text, range } => (text, Some(range)),
                    FormulaOut::Keep => bail!("a formula to keep is not rewritten"),
                };
                let mut f = BytesStart::new("f");
                if let Some(range) = range {
                    f.push_attribute(("t", "array"));
                    f.push_attribute(("ref", range.as_str()));
                }
                writer.write_event(Event::Start(f))?;
                writer.write_event(Event::Text(BytesText::new(text)))?;
                writer.write_event(Event::End(BytesEnd::new("f")))?;
                in_formula = matches!(event, Event::Start(_));
            }
            Event::End(e) if e.name().as_ref() == "f" => in_formula = false,
            _ if in_formula => {}
            Event::Start(e) if drop_value && e.name().as_ref() == "v" => in_value = true,
            Event::End(e) if drop_value && e.name().as_ref() == "v" => in_value = false,
            Event::Empty(e) if drop_value && e.name().as_ref() == "v" => {}
            _ if in_value => {}
            event => writer.write_event(event.borrow())?,
        }
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

#[cfg(test)]
mod tests {
    use super::*;

    /// `workbook.xml` after removing the sheets at `removed` and the names in `ranges`.
    fn modified_workbook(xml: &str, removed: &[usize], ranges: Option<&HashSet<String>>) -> String {
        let names: Vec<String> = parse_sheets(xml)
            .unwrap()
            .into_iter()
            .map(|sheet| sheet.name)
            .collect();
        let removed_names: HashSet<String> = removed.iter().map(|&i| names[i].clone()).collect();
        let plan = StructureEdits::new(Some(&names), &removed_names, &[], &[]).unwrap();
        let edits = plan.edits();
        modify_workbook_xml(
            xml,
            &WorkbookEdits {
                removed_indices: removed,
                removed_ranges: ranges,
                plan: &plan,
                edits: &edits,
                sheet_names: &names,
                recalculate_on_open: false,
            },
        )
        .unwrap()
    }

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
        let xml = modified_workbook(WORKBOOK_XML, &[0], None);
        assert!(!xml.contains("R&amp;D"));
        assert!(xml.contains(r#"<sheet name="Data" sheetId="2" r:id="rId2"/>"#));
    }

    #[test]
    fn test_remove_named_range_with_escaped_name() {
        let remove = HashSet::from(["A&B".to_string()]);
        let xml = modified_workbook(WORKBOOK_XML, &[], Some(&remove));
        assert!(!xml.contains("A&amp;B"));
        assert!(xml.contains(r#"<definedName name="Keep">Data!$B$1</definedName>"#));
    }

    #[test]
    fn test_remove_empty_defined_names() {
        let xml = r#"<workbook><sheets><sheet name="A"/><sheet name="B"/></sheets><definedNames><definedName name="Drop"/><definedName name="Local" localSheetId="0"/><definedName name="Moved" localSheetId="1"/></definedNames></workbook>"#;
        let remove = HashSet::from(["Drop".to_string()]);
        let modified = modified_workbook(xml, &[0], Some(&remove));
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
        let modified = modified_workbook(xml, &[1, 3], None);
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

    /// `workbook.xml` after `recalculate_on_open`, with no other change.
    fn recalculating(xml: &str) -> String {
        let names: Vec<String> = parse_sheets(xml)
            .unwrap()
            .into_iter()
            .map(|sheet| sheet.name)
            .collect();
        let plan = StructureEdits::new(Some(&names), &HashSet::new(), &[], &[]).unwrap();
        let edits = plan.edits();
        modify_workbook_xml(
            xml,
            &WorkbookEdits {
                removed_indices: &[],
                removed_ranges: None,
                plan: &plan,
                edits: &edits,
                sheet_names: &names,
                recalculate_on_open: true,
            },
        )
        .unwrap()
    }

    #[test]
    fn recalculation_on_open_finds_its_place_in_workbook_xml() {
        let sheets = r#"<sheets><sheet name="T"/></sheets>"#;
        // An existing `calcPr`, empty or with content: one element with the flag.
        for calc in [r#"<calcPr calcId="1"/>"#, r#"<calcPr calcId="1"></calcPr>"#] {
            let xml = recalculating(&format!("<workbook>{sheets}{calc}</workbook>"));
            assert_eq!(xml.matches("<calcPr").count(), 1, "{xml}");
            assert!(xml.contains(r#"fullCalcOnLoad="1""#), "{xml}");
        }
        // None: before the elements that follow it in the schema, or at the end.
        assert_eq!(
            recalculating(&format!("<workbook>{sheets}<extLst/></workbook>")),
            format!(r#"<workbook>{sheets}<calcPr fullCalcOnLoad="1"/><extLst/></workbook>"#)
        );
        assert_eq!(
            recalculating(&format!(
                "<workbook>{sheets}<pivotCaches></pivotCaches></workbook>"
            )),
            format!(
                r#"<workbook>{sheets}<calcPr fullCalcOnLoad="1"/><pivotCaches></pivotCaches></workbook>"#
            )
        );
        assert_eq!(
            recalculating(&format!("<workbook>{sheets}</workbook>")),
            format!(r#"<workbook>{sheets}<calcPr fullCalcOnLoad="1"/></workbook>"#)
        );
    }

    #[test]
    fn a_removed_name_is_not_taken_from_a_quoted_sheet_name() {
        let removed = HashSet::from(["Data".to_string()]);
        assert_eq!(
            replace_names("'Data Entry'!A1+Data+data", &removed),
            "'Data Entry'!A1+#NAME?+#NAME?"
        );
        assert_eq!(
            replace_names("\"Data\"&Data(1)", &removed),
            "\"Data\"&Data(1)"
        );
    }
}

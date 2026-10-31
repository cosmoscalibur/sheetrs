//! ODS writer functionality for removing sheets and named ranges, and for
//! converting formulas to values

use super::ods_structure::{
    adjust_range_list, move_end_anchors, rewrite_cell_formula, rewrite_expression, sheet_of_address,
};
use super::structure::{StructureEdits, shape_depends_on_sizes};
use super::xlsx_structure::{SheetStructure, map_attributes};
use super::xlsx_writer::replace_names;
use super::{
    ValueScope, ValueTarget, WorkbookModifications, attribute, read_element_body, write_element,
};
use anyhow::{Context, Result, bail};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashSet;
use std::fs::File;
use std::io::{BufRead, BufReader, BufWriter, Seek, Write};
use std::path::Path;
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

/// Part holding the sheets and named ranges of an ODS document.
const CONTENT_PART: &str = "content.xml";

const ROW: &str = "table:table-row";
const ROWS_REPEATED: &str = "table:number-rows-repeated";
const COLUMNS_REPEATED: &str = "table:number-columns-repeated";
const FORMULA: &str = "table:formula";
const MATRIX_COLUMNS: &str = "table:number-matrix-columns-spanned";
const MATRIX_ROWS: &str = "table:number-matrix-rows-spanned";

/// Modify the ODS file at `input` by applying specified modifications.
pub fn modify_workbook_ods<W: Write + Seek>(
    input: &Path,
    output: W,
    modifications: &WorkbookModifications,
) -> Result<()> {
    let mut archive = ZipArchive::new(File::open(input)?)?;
    let mut zip_writer = ZipWriter::new(output);

    for i in 0..archive.len() {
        if archive.by_index_raw(i)?.name() != CONTENT_PART {
            // Raw copies keep `mimetype` first and stored uncompressed, as
            // ODF requires.
            zip_writer.raw_copy_file(archive.by_index_raw(i)?)?;
            continue;
        }
        zip_writer.start_file(CONTENT_PART, SimpleFileOptions::default())?;
        modify_content_xml(
            BufReader::new(archive.by_index(i)?),
            &mut zip_writer,
            modifications,
            |plan| plan.with_order_source(input.to_path_buf(), sheet_names_in_file),
        )
        .with_context(|| format!("Failed to modify {CONTENT_PART}"))?;
    }

    zip_writer.finish()?;
    Ok(())
}

/// The names of the sheets of the ODS file at `path`, in order.
fn sheet_names_in_file(path: &Path) -> Result<Vec<String>> {
    let mut archive = ZipArchive::new(File::open(path)?)?;
    sheet_names(BufReader::new(archive.by_name(CONTENT_PART)?))
        .with_context(|| format!("Failed to read the sheets of {CONTENT_PART}"))
}

fn is_cell(name: &str) -> bool {
    matches!(name, "table:table-cell" | "table:covered-table-cell")
}

/// Elements the reader lists as defined names, so they can be removed by name.
fn is_named_range(name: &str) -> bool {
    matches!(
        name,
        "table:named-range" | "table:named-expression" | "table:database-range"
    )
}

fn repeat_count(element: &BytesStart, key: &str) -> Result<u32> {
    Ok(attribute(element, key)?
        .and_then(|n| n.parse::<u32>().ok())
        .unwrap_or(1)
        .max(1))
}

/// The names of the sheets of `content.xml`, in order.
fn sheet_names<R: BufRead>(input: R) -> Result<Vec<String>> {
    let mut reader = Reader::from_reader(input);
    let mut names = Vec::new();
    let (mut depth, mut buf) = (0u32, Vec::new());
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name().as_ref() == "table:table" => {
                depth += 1;
                if depth == 1 {
                    names.extend(attribute(&e, "table:name")?);
                }
            }
            Event::Empty(e) if depth == 0 && e.name().as_ref() == "table:table" => {
                names.extend(attribute(&e, "table:name")?);
            }
            Event::End(e) if e.name().as_ref() == "table:table" => depth -= 1,
            Event::Eof => return Ok(names),
            _ => {}
        }
    }
}

/// Apply the modifications to `content.xml`: drop removed sheets and named
/// ranges, and convert formulas to values in the sheets of `convert_to_values`.
fn modify_content_xml<R: BufRead, W: Write>(
    input: R,
    output: W,
    modifications: &WorkbookModifications,
    with_sheets: impl FnOnce(StructureEdits) -> StructureEdits,
) -> Result<()> {
    let removes_sheet = |name: &str| {
        modifications
            .remove_sheets
            .as_ref()
            .is_some_and(|sheets| sheets.contains(name))
    };
    // The sheets of the document are learned while reading: a scope on a sheet that is not found
    // is reported at the end.
    let removed_names: HashSet<String> = modifications.remove_sheets.clone().unwrap_or_default();
    let plan = StructureEdits::new(
        None,
        &removed_names,
        &modifications.delete_rows,
        &modifications.delete_columns,
    )?;
    let plan = with_sheets(plan);
    let edits = plan.edits();
    let removed_ranges = modifications
        .remove_named_ranges
        .as_ref()
        .filter(|ranges| !ranges.is_empty());
    // What names read: references to the sheets, with no sheet of their own.
    let names_view = SheetStructure {
        plan: &plan,
        edits: &edits,
        sheet: String::new(),
        rows: Default::default(),
        cols: Default::default(),
        recalculate: modifications.recalculate_on_open,
    };
    let mut structure: Option<SheetStructure> = None;
    let mut column_cursor = 0u32;
    // The shapes of the cells that go, waiting for the cell that takes their place.
    let mut orphans: Vec<Orphan> = Vec::new();

    let mut reader = Reader::from_reader(input);
    let mut writer = Writer::new(BufWriter::new(output));
    let mut kept_sheets = HashSet::new();
    let mut removed_sheets = 0usize;
    // Nesting of `table:table`: cells may hold tables of their own.
    let mut table_depth = 0u32;
    let mut scopes: Vec<&ValueScope> = Vec::new();
    let mut sheet = String::new();
    let mut row = 0u32;
    let mut buf = Vec::new();
    let mut skip_buf = Vec::new();

    loop {
        buf.clear();
        let event = reader.read_event_into(&mut buf)?;
        if let Event::Start(e) | Event::Empty(e) = &event
            && plan.deletes_rows_or_columns()
            && matches!(
                e.name().as_ref(),
                "table:data-pilot-table"
                    | "table:filter"
                    | "table:consolidation"
                    | "table:scenario"
            )
        {
            bail!(
                "the document has {} (`{}`), whose ranges are not adjusted yet",
                match e.name().as_ref() {
                    "table:data-pilot-table" => "pivot tables",
                    "table:filter" => "a filter",
                    "table:consolidation" => "a consolidation",
                    _ => "scenarios",
                },
                e.name().as_ref()
            );
        }
        match event {
            Event::Start(e) if e.name().as_ref() == "table:table" => {
                table_depth += 1;
                if table_depth == 1 {
                    let name = attribute(&e, "table:name")?.unwrap_or_default();
                    if removes_sheet(&name) {
                        let end = e.to_end().into_owned();
                        reader.read_to_end_into(end.name(), &mut skip_buf)?;
                        table_depth -= 1;
                        removed_sheets += 1;
                        continue;
                    }
                    scopes = modifications
                        .convert_to_values
                        .iter()
                        .filter(|scope| scope.sheet == name)
                        .collect();
                    kept_sheets.insert(name.clone());
                    structure =
                        (!plan.is_empty() || modifications.recalculate_on_open).then(|| {
                            SheetStructure {
                                plan: &plan,
                                edits: &edits,
                                sheet: name.clone(),
                                rows: plan.rows_of(&name),
                                cols: plan.cols_of(&name),
                                recalculate: modifications.recalculate_on_open,
                            }
                        });
                    column_cursor = 0;
                    orphans.clear();
                    sheet = name;
                    row = 0;
                    if let Some(s) = structure.as_ref().filter(|s| s.changes_layout()) {
                        let e = print_ranges(&e, s)?;
                        writer.write_event(Event::Start(e))?;
                        continue;
                    }
                }
                writer.write_event(Event::Start(e))?;
            }
            Event::End(e) if e.name().as_ref() == "table:table" => {
                table_depth = table_depth.saturating_sub(1);
                if table_depth == 0
                    && !orphans.is_empty()
                    && let Some(s) = structure.as_ref()
                {
                    write_orphan_rows(
                        &mut writer,
                        std::mem::take(&mut orphans),
                        s.rows.shifted(row),
                    )?;
                }
                writer.write_event(Event::End(e))?;
            }
            Event::Empty(e) if e.name().as_ref() == "table:table" => {
                if table_depth == 0 {
                    let name = attribute(&e, "table:name")?.unwrap_or_default();
                    if removes_sheet(&name) {
                        removed_sheets += 1;
                        continue;
                    }
                    kept_sheets.insert(name);
                }
                writer.write_event(Event::Empty(e))?;
            }
            Event::Start(e) if e.name().as_ref() == ROW && table_depth == 1 => {
                let repeated = repeat_count(&e, ROWS_REPEATED)?;
                if scopes.is_empty() && structure.is_none() {
                    writer.write_event(Event::Start(e))?;
                } else {
                    let body = read_element_body(&mut reader, ROW)?;
                    let mut ctx = RowWriting {
                        scopes: &scopes,
                        structure: structure.as_ref(),
                        orphans: &mut orphans,
                    };
                    write_rows(&mut writer, &e, &body, row, repeated, &mut ctx)
                        .with_context(|| format!("in sheet '{sheet}', row {}", row + 1))?;
                }
                row = row.saturating_add(repeated);
            }
            Event::Empty(e) if e.name().as_ref() == ROW && table_depth == 1 => {
                let repeated = repeat_count(&e, ROWS_REPEATED)?;
                match structure.as_ref().filter(|s| s.changes_layout()) {
                    Some(_) if !orphans.is_empty() => {
                        let mut ctx = RowWriting {
                            scopes: &scopes,
                            structure: structure.as_ref(),
                            orphans: &mut orphans,
                        };
                        write_rows(&mut writer, &e, &[], row, repeated, &mut ctx)
                            .with_context(|| format!("in sheet '{sheet}', row {}", row + 1))?;
                    }
                    Some(s) => {
                        let kept = s.rows.kept_in(row, repeated);
                        if kept > 0 {
                            writer.write_event(Event::Empty(rebuild_start(
                                &e,
                                ROWS_REPEATED,
                                kept,
                                |_| false,
                            )?))?;
                        }
                    }
                    None => writer.write_event(Event::Empty(e))?,
                }
                row = row.saturating_add(repeated);
            }
            Event::Start(e) if e.name().as_ref() == "table:table-column" && table_depth == 1 => {
                // A column definition has no content to keep.
                let end = e.to_end().into_owned();
                reader.read_to_end_into(end.name(), &mut skip_buf)?;
                write_column(&mut writer, e, structure.as_ref(), &mut column_cursor)?;
            }
            Event::Empty(e) if e.name().as_ref() == "table:table-column" && table_depth == 1 => {
                write_column(&mut writer, e, structure.as_ref(), &mut column_cursor)?;
            }
            Event::Start(e) if is_named_range(e.name().as_ref()) => {
                let name = attribute(&e, "table:name")?.unwrap_or_default();
                match defined_name_after(&e, &name, removed_ranges, &names_view, &plan)? {
                    Some(e) => writer.write_event(Event::Start(e))?,
                    None => {
                        let end = e.to_end().into_owned();
                        reader.read_to_end_into(end.name(), &mut skip_buf)?;
                    }
                }
            }
            Event::Empty(e) if is_named_range(e.name().as_ref()) => {
                let name = attribute(&e, "table:name")?.unwrap_or_default();
                if let Some(e) = defined_name_after(&e, &name, removed_ranges, &names_view, &plan)?
                {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Event::Start(e) if !plan.is_empty() && is_range_holder(e.name().as_ref()) => {
                let validation = validation_view(&e, &plan, &edits)?;
                let view = validation
                    .as_ref()
                    .or(structure.as_ref())
                    .unwrap_or(&names_view);
                match adjust_range_holder(&e, view)? {
                    Some(e) => writer.write_event(Event::Start(e))?,
                    None => {
                        let end = e.to_end().into_owned();
                        reader.read_to_end_into(end.name(), &mut skip_buf)?;
                    }
                }
            }
            Event::Empty(e) if !plan.is_empty() && is_range_holder(e.name().as_ref()) => {
                let validation = validation_view(&e, &plan, &edits)?;
                let view = validation
                    .as_ref()
                    .or(structure.as_ref())
                    .unwrap_or(&names_view);
                if let Some(e) = adjust_range_holder(&e, view)? {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Event::Eof => break,
            event => {
                // A table with a source is the cache of an external link,
                // which the reader does not list as a sheet.
                if let Event::Start(e) | Event::Empty(e) = &event
                    && table_depth == 1
                    && e.name().as_ref() == "table:table-source"
                {
                    kept_sheets.remove(&sheet);
                }
                writer.write_event(event)?;
            }
        }
    }

    if removed_sheets > 0 && kept_sheets.is_empty() {
        bail!("Cannot remove every sheet of the workbook");
    }
    for scope in modifications
        .convert_to_values
        .iter()
        .chain(&modifications.delete_rows)
        .chain(&modifications.delete_columns)
    {
        if !kept_sheets.contains(&scope.sheet) && !removes_sheet(&scope.sheet) {
            bail!("Sheet '{}' not found in {CONTENT_PART}", scope.sheet);
        }
    }

    if let Some(error) = plan.order_error() {
        bail!("cannot read the sheets of the document: {error}");
    }
    writer.into_inner().flush()?;
    Ok(())
}

/// Split `first..first + count` where a band of `bands` starts or ends, so
/// that every part lies wholly inside or outside each band.
fn split_repeats(
    first: u32,
    count: u32,
    bands: impl Iterator<Item = (u32, u32)>,
) -> Vec<(u32, u32)> {
    let end = first.saturating_add(count);
    let mut cuts: Vec<u32> = bands
        .flat_map(|(start, last)| [start, last.saturating_add(1)])
        .filter(|&cut| cut > first && cut < end)
        .collect();
    cuts.push(end);
    cuts.sort_unstable();
    cuts.dedup();

    let mut parts = Vec::with_capacity(cuts.len());
    let mut start = first;
    for cut in cuts {
        parts.push((start, cut - start));
        start = cut;
    }
    parts
}

/// Copy `element` with `repeat_key` set to `count`, dropping the attributes
/// for which `drop` holds.
fn rebuild_start<'a>(
    element: &BytesStart<'a>,
    repeat_key: &str,
    count: u32,
    drop: impl Fn(&str) -> bool,
) -> Result<BytesStart<'a>> {
    let mut rebuilt = element.clone();
    rebuilt.clear_attributes();
    for attr in element.attributes() {
        let attr = attr?;
        let key = attr.key.as_ref();
        if key != repeat_key && !drop(key) {
            rebuilt.push_attribute(attr);
        }
    }
    if count > 1 {
        rebuilt.push_attribute((repeat_key, count.to_string().as_str()));
    }
    Ok(rebuilt)
}

/// A direct child of a row: a cell with its content, or any other event.
enum RowItem<'b, 'a> {
    Cell {
        start: &'b BytesStart<'a>,
        content: Option<&'b [Event<'a>]>,
    },
    Other(&'b Event<'a>),
}

fn row_items<'b, 'a>(body: &'b [Event<'a>]) -> Vec<RowItem<'b, 'a>> {
    let mut items = Vec::new();
    let mut i = 0;
    while i < body.len() {
        match &body[i] {
            Event::Empty(e) if is_cell(e.name().as_ref()) => {
                items.push(RowItem::Cell {
                    start: e,
                    content: None,
                });
                i += 1;
            }
            Event::Start(e) if is_cell(e.name().as_ref()) => {
                let mut depth = 0u32;
                let mut end = i + 1;
                while end < body.len() {
                    match &body[end] {
                        Event::Start(_) => depth += 1,
                        Event::End(_) if depth == 0 => break,
                        Event::End(_) => depth -= 1,
                        _ => {}
                    }
                    end += 1;
                }
                items.push(RowItem::Cell {
                    start: e,
                    content: Some(&body[i + 1..end]),
                });
                i = end + 1;
            }
            event => {
                items.push(RowItem::Other(event));
                i += 1;
            }
        }
    }
    items
}

/// A shape that was in a cell that goes. LibreOffice keeps it where it was, anchored to the cell
/// that is there once the rows and columns are gone: the one with the index of the cell it left
/// (the rows and columns that pass have the same height and width in a document that does not
/// say otherwise, and a size is not in the file).
struct Orphan {
    /// The row and column the shape is anchored to now.
    row: u32,
    col: u32,
    shapes: Vec<Event<'static>>,
}

/// What rows and cells are written under.
struct RowWriting<'a, 'b> {
    /// The scopes whose formulas become values.
    scopes: &'a [&'a ValueScope],
    structure: Option<&'a SheetStructure<'b>>,
    /// The shapes of the cells that go, waiting for the cell that takes their place.
    orphans: &'a mut Vec<Orphan>,
}

/// The shapes of the cell at `(row, col)` of the input, which a deletion removes, as an orphan.
fn orphan_of(
    s: &SheetStructure,
    row: u32,
    col: u32,
    shapes: Vec<Event<'static>>,
) -> Result<Orphan> {
    if s.rows.deleted_below_a_deleted(row) || s.cols.deleted_below_a_deleted(col) {
        return Err(shape_depends_on_sizes(&s.sheet));
    }
    Ok(Orphan {
        row: s.rows.anchored_at(row),
        col: s.cols.anchored_at(col),
        shapes: move_end_anchors(&shapes, s),
    })
}

/// Write a `table:table-row` covering `repeated` rows from `first_row`, converting the formulas of
/// its cells that fall inside the scopes of `ctx` and applying its structure: the rows and columns
/// it deletes go, and the formulas that read them change. The shapes of the cells that go are
/// written into the cell that takes their place.
///
/// Repeated rows and cells are split where a scope starts or ends, or a shape is written, so each
/// written element lies wholly inside or outside every scope.
fn write_rows<W: Write>(
    writer: &mut Writer<W>,
    start: &BytesStart,
    body: &[Event],
    first_row: u32,
    repeated: u32,
    ctx: &mut RowWriting,
) -> Result<()> {
    let items = row_items(body);
    let layout = ctx.structure.filter(|s| s.changes_layout());
    let mut has_formula = false;
    for item in &items {
        if let RowItem::Cell { start, .. } = item
            && start.try_get_attribute(FORMULA)?.is_some()
        {
            has_formula = true;
            break;
        }
    }
    if !has_formula && layout.is_none() {
        writer.write_event(Event::Start(start.borrow()))?;
        for event in body {
            writer.write_event(event.borrow())?;
        }
        writer.write_event(Event::End(start.to_end()))?;
        return Ok(());
    }

    let row_bands = ctx.scopes.iter().filter_map(|scope| match scope.target {
        ValueTarget::Rows(first, last) => Some((first, last)),
        _ => None,
    });
    for (row, count) in split_repeats(first_row, repeated, row_bands) {
        let kept = layout.map_or(count, |s| s.rows.kept_in(row, count));
        if let Some(s) = layout
            && kept == 0
        {
            let mut col = 0u32;
            for item in &items {
                if let RowItem::Cell { start, content } = item {
                    check_merged_area(ctx.structure, start, row, col)?;
                    let shapes = content.map(shapes_of).unwrap_or_default();
                    if !shapes.is_empty() {
                        ctx.orphans.push(orphan_of(s, row, col, shapes)?);
                    }
                    col = col.saturating_add(repeat_count(start, COLUMNS_REPEATED)?);
                }
            }
            continue;
        }
        // The rows that stay in their new place: one by one where a shape waits for its row.
        let new_first = layout.map_or(row, |s| s.rows.shifted(row));
        let new_end = new_first + kept;
        let mut waiting: Vec<u32> = ctx
            .orphans
            .iter()
            .map(|o| o.row)
            .filter(|r| (new_first..new_end).contains(r))
            .collect();
        waiting.sort_unstable();
        waiting.dedup();
        let mut next = new_first;
        let mut collects = true;
        for target in waiting {
            if target > next {
                write_row(
                    writer,
                    start,
                    &items,
                    (row, next),
                    target - next,
                    collects,
                    ctx,
                )?;
                collects = false;
            }
            write_row(writer, start, &items, (row, target), 1, collects, ctx)?;
            collects = false;
            next = target + 1;
        }
        if next < new_end {
            write_row(
                writer,
                start,
                &items,
                (row, next),
                new_end - next,
                collects,
                ctx,
            )?;
        }
    }
    Ok(())
}

/// Write `count` rows of the element `start` that sit at `new_row` of the output, the first of
/// them coming from `row` of the input.
fn write_row<W: Write>(
    writer: &mut Writer<W>,
    start: &BytesStart,
    items: &[RowItem],
    (row, new_row): (u32, u32),
    count: u32,
    collects: bool,
    ctx: &mut RowWriting,
) -> Result<()> {
    writer.write_event(Event::Start(rebuild_start(
        start,
        ROWS_REPEATED,
        count,
        |_| false,
    )?))?;
    let (mut col, mut written) = (0u32, 0u32);
    for item in items {
        match item {
            RowItem::Cell { start, content } => {
                let repeated = repeat_count(start, COLUMNS_REPEATED)?;
                let run = CellRun {
                    row,
                    new_row,
                    first_col: col,
                    repeated,
                    collects,
                };
                written += write_cells(writer, start, *content, run, ctx)?;
                col = col.saturating_add(repeated);
            }
            RowItem::Other(event) => writer.write_event(event.borrow())?,
        }
    }
    let (here, other) = std::mem::take(ctx.orphans)
        .into_iter()
        .partition(|o| o.row == new_row);
    *ctx.orphans = other;
    write_orphan_cells(writer, here, written)?;
    writer.write_event(Event::End(start.to_end()))?;
    Ok(())
}

/// Write the cells that hold the shapes of `orphans`, all of one row, after the `written` columns
/// that the row has.
fn write_orphan_cells<W: Write>(
    writer: &mut Writer<W>,
    mut orphans: Vec<Orphan>,
    written: u32,
) -> Result<()> {
    orphans.sort_by_key(|o| o.col);
    let mut next = written;
    let mut orphans = orphans.into_iter().peekable();
    while let Some(first) = orphans.next() {
        let mut shapes = first.shapes;
        while let Some(same) = orphans.next_if(|o| o.col == first.col) {
            shapes.extend(same.shapes);
        }
        if first.col > next {
            let mut gap = BytesStart::new("table:table-cell");
            if first.col - next > 1 {
                gap.push_attribute((COLUMNS_REPEATED, (first.col - next).to_string().as_str()));
            }
            writer.write_event(Event::Empty(gap))?;
        }
        write_element(writer, BytesStart::new("table:table-cell"), Some(&shapes))?;
        next = first.col + 1;
    }
    Ok(())
}

/// Write the rows that hold the shapes of `orphans` that no row of the output has, in the
/// rows they are anchored to, after the `next` row that the table would have next; empty rows
/// fill the gaps.
fn write_orphan_rows<W: Write>(
    writer: &mut Writer<W>,
    mut orphans: Vec<Orphan>,
    mut next: u32,
) -> Result<()> {
    orphans.sort_by_key(|o| (o.row, o.col));
    while let Some(row) = orphans.first().map(|o| o.row) {
        let at = orphans.partition_point(|o| o.row == row);
        let here: Vec<Orphan> = orphans.drain(..at).collect();
        if row > next {
            let mut gap = BytesStart::new(ROW);
            if row - next > 1 {
                gap.push_attribute((ROWS_REPEATED, (row - next).to_string().as_str()));
            }
            write_element(
                writer,
                gap,
                Some(&[Event::Empty(BytesStart::new("table:table-cell"))]),
            )?;
        }
        writer.write_event(Event::Start(BytesStart::new(ROW)))?;
        write_orphan_cells(writer, here, 0)?;
        writer.write_event(Event::End(BytesEnd::new(ROW)))?;
        next = row + 1;
    }
    Ok(())
}

/// Where a cell element sits: the row of the run it is in and the columns it covers.
struct CellRun {
    /// The row of the input, and the one it has in the output.
    row: u32,
    new_row: u32,
    first_col: u32,
    repeated: u32,
    /// Whether the shapes of its deleted cells are collected: once for a run of repeated rows that
    /// is written in several pieces.
    collects: bool,
}

/// The cached result of a formula cell: its type and value attributes, and its displayed text.
const CACHED_VALUE_ATTRIBUTES: [&str; 8] = [
    "office:value-type",
    "office:value",
    "office:string-value",
    "office:boolean-value",
    "office:date-value",
    "office:time-value",
    "office:currency",
    "calcext:value-type",
];

/// The events of a cell without its text paragraphs, which show the cached result; annotations
/// and other content stay.
fn without_paragraphs<'a>(content: &[Event<'a>]) -> Vec<Event<'a>> {
    let mut kept = Vec::new();
    // Elements open inside the cell, and the depth inside a paragraph that is being dropped.
    let (mut nest, mut dropped) = (0u32, 0u32);
    for event in content {
        match event {
            Event::Start(_) if dropped > 0 => dropped += 1,
            Event::Start(e) if nest == 0 && e.name().as_ref() == "text:p" => dropped = 1,
            Event::End(_) if dropped > 0 => dropped -= 1,
            Event::Empty(e) if dropped > 0 || (nest == 0 && e.name().as_ref() == "text:p") => {}
            _ if dropped > 0 => {}
            Event::Start(_) => {
                nest += 1;
                kept.push(event.clone());
            }
            Event::End(_) => {
                nest = nest.saturating_sub(1);
                kept.push(event.clone());
            }
            event => kept.push(event.clone()),
        }
    }
    kept
}

/// The shapes a cell holds (a drawing, an image, a chart): the elements `draw:*` of its content.
/// A cell carries the shapes that are anchored to it.
fn shapes_of(events: &[Event]) -> Vec<Event<'static>> {
    let is_shape = |e: &BytesStart| e.name().as_ref().starts_with("draw:");
    let (mut shapes, mut depth, mut inside) = (Vec::new(), 0u32, false);
    for event in events {
        match event {
            Event::Start(e) => {
                inside |= depth == 0 && is_shape(e);
                depth += 1;
                if inside {
                    shapes.push(event.clone().into_owned());
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if inside {
                    shapes.push(event.clone().into_owned());
                    inside = depth > 0;
                }
            }
            Event::Empty(e) if depth == 0 && is_shape(e) => {
                shapes.push(event.clone().into_owned());
            }
            _ if inside => shapes.push(event.clone().into_owned()),
            _ => {}
        }
    }
    shapes
}

/// The content of a cell with `shapes` added after its note, which a cell lists first.
fn with_shapes<'a>(content: Option<&[Event<'a>]>, shapes: &[Event<'static>]) -> Vec<Event<'a>> {
    let content = content.unwrap_or_default();
    let mut after_note = 0;
    if let Some(Event::Start(e)) = content.first()
        && e.name().as_ref() == "office:annotation"
    {
        let mut depth = 0u32;
        for (i, event) in content.iter().enumerate() {
            match event {
                Event::Start(_) => depth += 1,
                Event::End(_) => depth -= 1,
                _ => {}
            }
            if depth == 0 {
                after_note = i + 1;
                break;
            }
        }
    }
    let mut out = content[..after_note].to_vec();
    out.extend(shapes.iter().cloned());
    out.extend_from_slice(&content[after_note..]);
    out
}

/// Write the `kept` columns of the cell `cell` that sit at `new_col` of `new_row`, with the shapes
/// that wait for one of them added to it; the columns before and after it stay as they were.
fn write_cell_run<W: Write>(
    writer: &mut Writer<W>,
    cell: BytesStart,
    content: Option<&[Event]>,
    (new_row, new_col): (u32, u32),
    kept: u32,
    orphans: &mut Vec<Orphan>,
) -> Result<()> {
    let columns = new_col..new_col + kept;
    let (mut here, other) = std::mem::take(orphans)
        .into_iter()
        .partition::<Vec<_>, _>(|o| o.row == new_row && columns.contains(&o.col));
    *orphans = other;
    here.sort_by_key(|o| o.col);
    let mut next = new_col;
    let mut here = here.into_iter().peekable();
    while let Some(first) = here.next() {
        let mut shapes = first.shapes;
        while let Some(same) = here.next_if(|o| o.col == first.col) {
            shapes.extend(same.shapes);
        }
        if first.col > next {
            let before = rebuild_start(&cell, COLUMNS_REPEATED, first.col - next, |_| false)?;
            write_element(writer, before, content)?;
        }
        let one = rebuild_start(&cell, COLUMNS_REPEATED, 1, |_| false)?;
        write_element(writer, one, Some(&with_shapes(content, &shapes)))?;
        next = first.col + 1;
    }
    if next == new_col {
        return write_element(writer, cell, content);
    }
    if next < columns.end {
        let after = rebuild_start(&cell, COLUMNS_REPEATED, columns.end - next, |_| false)?;
        write_element(writer, after, content)?;
    }
    Ok(())
}

/// A merged area whose first cell is deleted while some of its cells stay has no first cell left.
fn check_merged_area(
    structure: Option<&SheetStructure>,
    cell: &BytesStart,
    row: u32,
    col: u32,
) -> Result<()> {
    let Some(s) = structure.filter(|s| s.changes_layout()) else {
        return Ok(());
    };
    let cols = repeat_count(cell, "table:number-columns-spanned")?;
    let rows = repeat_count(cell, "table:number-rows-spanned")?;
    if (cols > 1 || rows > 1) && s.rows.kept_in(row, rows) > 0 && s.cols.kept_in(col, cols) > 0 {
        bail!(
            "a deletion in sheet '{}' removes the first cell of a merged area that stays",
            s.sheet
        );
    }
    Ok(())
}

/// Write a cell covering `repeated` columns from `first_col` of `row`: the formulas that fall
/// inside `scopes` become values, the columns that `structure` deletes go, and a formula that
/// reads what the plan changes is rewritten (it loses its cached result, no longer valid).
fn write_cells<W: Write>(
    writer: &mut Writer<W>,
    start: &BytesStart,
    content: Option<&[Event]>,
    run: CellRun,
    ctx: &mut RowWriting,
) -> Result<u32> {
    let CellRun {
        row,
        new_row,
        first_col,
        repeated,
        collects,
    } = run;
    let layout = ctx.structure.filter(|s| s.changes_layout());
    let formula = attribute(start, FORMULA)?;
    // LibreOffice stores an empty text result as a paragraph without a value
    // type. A formula with neither has no cached value to convert to.
    let value_type = attribute(start, "office:value-type")?;
    let has_paragraph = content.is_some_and(|events| {
        events.iter().any(|event| {
            matches!(event, Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "text:p")
        })
    });
    let convertible = formula.is_some() && !(value_type.is_none() && !has_paragraph);
    let new_formula = match (ctx.structure, formula.as_deref()) {
        (Some(s), Some(f)) => match rewrite_cell_formula(s, f)? {
            None if s.recalculate => Some(f.to_string()),
            rewritten => rewritten,
        },
        _ => None,
    };
    if !convertible && layout.is_none() && new_formula.is_none() {
        write_element(writer, start.clone(), content)?;
        return Ok(repeated);
    }

    let is_error = attribute(start, "calcext:value-type")?.as_deref() == Some("error");
    let matrix_cols = repeat_count(start, MATRIX_COLUMNS)?;
    let matrix_rows = repeat_count(start, MATRIX_ROWS)?;
    let column_bands: Vec<(u32, u32)> = if convertible {
        ctx.scopes
            .iter()
            .filter_map(|scope| match scope.target {
                ValueTarget::Columns(first, last) => Some((first, last)),
                _ => None,
            })
            .collect()
    } else {
        Vec::new()
    };
    let mut written = 0u32;
    for (col, count) in split_repeats(first_col, repeated, column_bands.into_iter()) {
        let kept = layout.map_or(count, |s| s.cols.kept_in(col, count));
        if let Some(s) = layout
            && kept == 0
        {
            check_merged_area(ctx.structure, start, row, col)?;
            let shapes = content.map(shapes_of).unwrap_or_default();
            if collects && !shapes.is_empty() {
                ctx.orphans.push(orphan_of(s, row, col, shapes)?);
            }
            continue;
        }
        let new_col = layout.map_or(col, |s| s.cols.shifted(col));
        if let Some(s) = layout
            && (matrix_cols > 1 || matrix_rows > 1)
            && (s.cols.kept_in(col, matrix_cols) != matrix_cols
                || s.rows.kept_in(row, matrix_rows) != matrix_rows)
        {
            bail!(
                "a deletion in sheet '{}' covers part of an array formula",
                s.sheet
            );
        }
        // Array formulas convert only when one scope covers their whole range,
        // since their other cells hold no formula of their own.
        let last = (row + matrix_rows - 1, col + matrix_cols - 1);
        let converts = convertible
            && ctx
                .scopes
                .iter()
                .any(|scope| scope.contains_range((row, col), last));
        let mut cell = if converts {
            let mut cell = rebuild_start(start, COLUMNS_REPEATED, kept, |key| {
                matches!(key, FORMULA | MATRIX_COLUMNS | MATRIX_ROWS)
                    // Error results show their text, such as `#DIV/0!`.
                    || (is_error && key == "office:string-value")
            })?;
            if value_type.is_none() {
                cell.push_attribute(("office:value-type", "string"));
                cell.push_attribute(("calcext:value-type", "string"));
            }
            cell
        } else {
            rebuild_start(start, COLUMNS_REPEATED, kept, |_| false)?
        };
        let mut shown: Option<Vec<Event>> = None;
        if !converts && let Some(new) = &new_formula {
            cell = map_attributes(&cell, |key, _| {
                Ok(if key == FORMULA {
                    Some(Some(new.clone()))
                } else if CACHED_VALUE_ATTRIBUTES.contains(&key) {
                    Some(None)
                } else {
                    None
                })
            })?
            .into_owned();
            shown = content.map(without_paragraphs);
        }
        if let Some(s) = layout {
            let cols = repeat_count(start, "table:number-columns-spanned")?;
            let rows = repeat_count(start, "table:number-rows-spanned")?;
            if cols > 1 || rows > 1 {
                let (new_cols, new_rows) = (s.cols.kept_in(col, cols), s.rows.kept_in(row, rows));
                cell = map_attributes(&cell, |key, _| {
                    Ok(match key {
                        "table:number-columns-spanned" => {
                            Some((new_cols > 1).then(|| new_cols.to_string()))
                        }
                        "table:number-rows-spanned" => {
                            Some((new_rows > 1).then(|| new_rows.to_string()))
                        }
                        _ => None,
                    })
                })?;
            }
        }
        let body = shown.as_deref().or(content);
        let moved;
        let body = match (body, layout) {
            (Some(events), Some(s)) if has_end_anchor(events) => {
                moved = move_end_anchors(events, s);
                Some(&moved[..])
            }
            (body, _) => body,
        };
        write_cell_run(writer, cell, body, (new_row, new_col), kept, ctx.orphans)?;
        written += kept;
    }
    Ok(written)
}

/// Whether a shape in `events` ends at a cell of its own (`table:end-cell-address`).
fn has_end_anchor(events: &[Event]) -> bool {
    events.iter().any(|event| {
        matches!(event, Event::Start(e) | Event::Empty(e)
            if matches!(e.try_get_attribute("table:end-cell-address"), Ok(Some(_))))
    })
}

/// A column definition: its run of columns without those the plan deletes.
fn write_column<W: Write>(
    writer: &mut Writer<W>,
    column: BytesStart,
    structure: Option<&SheetStructure>,
    cursor: &mut u32,
) -> Result<()> {
    let repeated = repeat_count(&column, COLUMNS_REPEATED)?;
    match structure.filter(|s| s.changes_layout()) {
        Some(s) => {
            let kept = s.cols.kept_in(*cursor, repeated);
            if kept > 0 {
                writer.write_event(Event::Empty(rebuild_start(
                    &column,
                    COLUMNS_REPEATED,
                    kept,
                    |_| false,
                )?))?;
            }
        }
        None => writer.write_event(Event::Empty(column))?,
    }
    *cursor = cursor.saturating_add(repeated);
    Ok(())
}

/// The `table:table` element of a sheet with its print ranges adjusted.
fn print_ranges(table: &BytesStart, s: &SheetStructure) -> Result<BytesStart<'static>> {
    map_attributes(table, |key, value| {
        Ok((key == "table:print-ranges").then(|| adjust_range_list(s, value)))
    })
}

/// A defined name (named range, named expression, database range) once the plan is applied, or
/// `None` when it goes: a removed name, or a database range with no cell left.
fn defined_name_after(
    element: &BytesStart,
    name: &str,
    removed_ranges: Option<&HashSet<String>>,
    view: &SheetStructure,
    plan: &StructureEdits,
) -> Result<Option<BytesStart<'static>>> {
    if removed_ranges.is_some_and(|ranges| ranges.contains(name)) {
        return Ok(None);
    }
    if plan.is_empty() && removed_ranges.is_none() {
        return Ok(Some(element.clone().into_owned()));
    }
    match element.name().as_ref() {
        // What LibreOffice writes for a range that is gone.
        "table:named-range" => Ok(Some(map_attributes(element, |key, value| {
            Ok((key == "table:cell-range-address").then(|| {
                Some(adjust_range_list(view, value).unwrap_or_else(|| "#REF!".to_string()))
            }))
        })?)),
        "table:named-expression" => Ok(Some(map_attributes(element, |key, value| {
            if key != "table:expression" {
                return Ok(None);
            }
            let adjusted = rewrite_expression(view, value).unwrap_or_else(|| value.to_string());
            let new = match removed_ranges {
                Some(removed) => replace_names(&adjusted, removed),
                None => adjusted,
            };
            Ok((new != value).then_some(Some(new)))
        })?)),
        "table:database-range" => {
            let range = attribute(element, "table:target-range-address")?.unwrap_or_default();
            match adjust_range_list(view, &range) {
                Some(new) => Ok(Some(map_attributes(element, |key, _| {
                    Ok((key == "table:target-range-address").then(|| Some(new.clone())))
                })?)),
                None => Ok(None),
            }
        }
        _ => Ok(Some(element.clone().into_owned())),
    }
}

/// The sheet a content validation reads its unqualified references on: the one of its base cell.
fn validation_view<'a>(
    element: &BytesStart,
    plan: &'a StructureEdits,
    edits: &'a [sheetrs_formula::edit::Edit<'a>],
) -> Result<Option<SheetStructure<'a>>> {
    if element.name().as_ref() != "table:content-validation" {
        return Ok(None);
    }
    let base = attribute(element, "table:base-cell-address")?.unwrap_or_default();
    Ok(sheet_of_address(&base).map(|sheet| SheetStructure {
        plan,
        edits,
        rows: plan.rows_of(&sheet),
        cols: plan.cols_of(&sheet),
        sheet,
        recalculate: false,
    }))
}

/// The elements that hold a range or an expression of the sheet outside its cells.
fn is_range_holder(name: &str) -> bool {
    matches!(
        name,
        "calcext:conditional-format" | "calcext:condition" | "table:content-validation"
    )
}

/// A range holder once the plan is applied, `None` when its range has no cell left.
fn adjust_range_holder(
    element: &BytesStart,
    view: &SheetStructure,
) -> Result<Option<BytesStart<'static>>> {
    match element.name().as_ref() {
        "calcext:conditional-format" => {
            let range = attribute(element, "calcext:target-range-address")?.unwrap_or_default();
            match adjust_range_list(view, &range) {
                Some(new) => Ok(Some(map_attributes(element, |key, _| {
                    Ok((key == "calcext:target-range-address").then(|| Some(new.clone())))
                })?)),
                None => Ok(None),
            }
        }
        name => {
            let key_of = if name == "calcext:condition" {
                "calcext:value"
            } else {
                "table:condition"
            };
            Ok(Some(map_attributes(element, |key, value| {
                Ok((key == key_of)
                    .then(|| rewrite_expression(view, value))
                    .flatten()
                    .map(Some))
            })?))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn scopes(specs: &[&str]) -> Vec<ValueScope> {
        specs.iter().map(|spec| spec.parse().unwrap()).collect()
    }

    fn convert(xml: &str, specs: &[&str]) -> String {
        let modifications = WorkbookModifications {
            convert_to_values: scopes(specs),
            ..Default::default()
        };
        modify_xml(xml, &modifications).unwrap()
    }

    fn modify_xml(xml: &str, modifications: &WorkbookModifications) -> Result<String> {
        let mut output = Vec::new();
        modify_content_xml(xml.as_bytes(), &mut output, modifications, |plan| {
            plan.with_order(sheet_names(xml.as_bytes()).unwrap())
        })?;
        Ok(String::from_utf8(output)?)
    }

    #[test]
    fn test_split_repeats() {
        assert_eq!(split_repeats(0, 5, [].into_iter()), vec![(0, 5)]);
        assert_eq!(
            split_repeats(0, 5, [(2, 3)].into_iter()),
            vec![(0, 2), (2, 2), (4, 1)]
        );
        assert_eq!(split_repeats(3, 2, [(0, 10)].into_iter()), vec![(3, 2)]);
        assert_eq!(
            split_repeats(3, 4, [(5, 5), (4, 6)].into_iter()),
            vec![(3, 1), (4, 1), (5, 1), (6, 1)]
        );
    }

    const SHEET: &str = r#"<table:table table:name="Data"><table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell><table:table-cell table:formula="of:=[.A1]*2" office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell></table:table-row></table:table>"#;

    #[test]
    fn test_convert_drops_formula_and_keeps_value() {
        let xml = convert(SHEET, &["Data"]);
        assert!(!xml.contains("table:formula"), "{xml}");
        assert!(xml.contains(r#"<table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell>"#), "{xml}");
    }

    #[test]
    fn test_convert_splits_repeated_cells_and_rows() {
        let xml = r#"<table:table table:name="Data"><table:table-row table:number-rows-repeated="3"><table:table-cell table:number-columns-repeated="3" table:formula="of:=1+1" office:value-type="float" office:value="2"/></table:table-row></table:table>"#;
        let converted = convert(xml, &["Data!2", "Data!B"]);
        let expected = concat!(
            r#"<table:table table:name="Data">"#,
            r#"<table:table-row><table:table-cell table:formula="of:=1+1" office:value-type="float" office:value="2"/><table:table-cell office:value-type="float" office:value="2"/><table:table-cell table:formula="of:=1+1" office:value-type="float" office:value="2"/></table:table-row>"#,
            r#"<table:table-row><table:table-cell office:value-type="float" office:value="2"/><table:table-cell office:value-type="float" office:value="2"/><table:table-cell office:value-type="float" office:value="2"/></table:table-row>"#,
            r#"<table:table-row><table:table-cell table:formula="of:=1+1" office:value-type="float" office:value="2"/><table:table-cell office:value-type="float" office:value="2"/><table:table-cell table:formula="of:=1+1" office:value-type="float" office:value="2"/></table:table-row>"#,
            r#"</table:table>"#
        );
        assert_eq!(converted, expected);
    }

    #[test]
    fn test_convert_error_and_uncached_formulas() {
        let xml = r#"<table:table table:name="Data"><table:table-row><table:table-cell table:formula="of:=1/0" office:value-type="string" office:string-value="" calcext:value-type="error"><text:p>#DIV/0!</text:p></table:table-cell><table:table-cell table:formula="of:=1+1"/><table:table-cell table:formula="of:=&quot;&quot;"><text:p/></table:table-cell></table:table-row></table:table>"#;
        let converted = convert(xml, &["Data"]);
        assert!(converted.contains(r#"<table:table-cell office:value-type="string" calcext:value-type="string"><text:p/></table:table-cell>"#), "{converted}");
        assert!(converted.contains(r#"<table:table-cell office:value-type="string" calcext:value-type="error"><text:p>#DIV/0!</text:p></table:table-cell>"#), "{converted}");
        assert!(
            converted.contains(r#"<table:table-cell table:formula="of:=1+1"/>"#),
            "{converted}"
        );
    }

    #[test]
    fn test_convert_array_formula_only_when_covered() {
        let xml = r#"<table:table table:name="Data"><table:table-row><table:table-cell table:number-matrix-columns-spanned="2" table:number-matrix-rows-spanned="1" table:formula="of:={1;2}" office:value-type="float" office:value="1"/><table:table-cell office:value-type="float" office:value="2"/></table:table-row></table:table>"#;
        assert_eq!(convert(xml, &["Data!A"]), xml);
        let converted = convert(xml, &["Data!A:B"]);
        assert!(!converted.contains("matrix"), "{converted}");
        assert!(!converted.contains("table:formula"), "{converted}");
    }

    #[test]
    fn test_other_sheets_and_nested_tables_are_untouched() {
        let other = r#"<table:table table:name="Other"><table:table-row><table:table-cell table:formula="of:=1" office:value-type="float" office:value="1"/></table:table-row></table:table>"#;
        let nested = SHEET.replace("<text:p>1</text:p>", other);
        let converted = convert(&format!("{other}{nested}"), &["Data"]);
        assert_eq!(
            converted.matches(r#"table:formula="of:=1""#).count(),
            2,
            "{converted}"
        );
        assert!(!converted.contains("of:=[.A1]*2"), "{converted}");
    }

    #[test]
    fn test_unknown_sheet_fails() {
        let modifications = WorkbookModifications {
            convert_to_values: scopes(&["Missing"]),
            ..Default::default()
        };
        let err = modify_xml(SHEET, &modifications).unwrap_err();
        assert!(err.to_string().contains("Missing"), "{err}");
    }

    #[test]
    fn test_remove_sheet_and_named_ranges() {
        let xml = format!(
            r#"<office:spreadsheet><table:table table:name="Old"><table:table-row/></table:table>{SHEET}<table:named-expressions><table:named-range table:name="Drop" table:base-cell-address="$Data.$A$1" table:cell-range-address="$Data.$A$1"/><table:named-expression table:name="Keep" table:expression="of:=1"/></table:named-expressions><table:database-ranges><table:database-range table:name="DropDb" table:target-range-address="Data.A1:Data.A2"/></table:database-ranges></office:spreadsheet>"#
        );
        let modifications = WorkbookModifications {
            remove_sheets: Some(HashSet::from(["Old".to_string()])),
            remove_named_ranges: Some(HashSet::from(["Drop".to_string(), "DropDb".to_string()])),
            ..Default::default()
        };
        let modified = modify_xml(&xml, &modifications).unwrap();
        assert!(!modified.contains("Old"), "{modified}");
        assert!(!modified.contains("Drop"), "{modified}");
        assert!(modified.contains(SHEET), "{modified}");
        assert!(modified.contains(r#"table:name="Keep""#), "{modified}");
    }

    fn structure(xml: &str, rows: &[&str], cols: &[&str]) -> Result<String> {
        modify_xml(
            xml,
            &WorkbookModifications {
                delete_rows: scopes(rows),
                delete_columns: scopes(cols),
                ..Default::default()
            },
        )
    }

    const DOC_HEAD: &str = "<office:document-content xmlns:office=\"o\" xmlns:table=\"t\" xmlns:text=\"x\" xmlns:calcext=\"c\"><office:body><office:spreadsheet>";
    const DOC_TAIL: &str = "</office:spreadsheet></office:body></office:document-content>";

    fn doc(inner: &str) -> String {
        format!("{DOC_HEAD}{inner}{DOC_TAIL}")
    }

    #[test]
    fn deleted_rows_shorten_the_runs_of_repeated_rows() {
        // Rows 1 (once), 2-6 (repeated five times), 7 (once).
        let xml = doc(
            r#"<table:table table:name="S"><table:table-row><table:table-cell><text:p>a</text:p></table:table-cell></table:table-row><table:table-row table:number-rows-repeated="5"><table:table-cell/></table:table-row><table:table-row><table:table-cell><text:p>z</text:p></table:table-cell></table:table-row></table:table>"#,
        );
        // Rows 4-5 go: three of the five stay.
        let out = structure(&xml, &["S!4:5"], &[]).unwrap();
        assert!(
            out.contains(r#"<table:table-row table:number-rows-repeated="3">"#),
            "{out}"
        );
        // The whole run goes, and one row of it.
        let out = structure(&xml, &["S!2:6"], &[]).unwrap();
        assert!(!out.contains("number-rows-repeated"), "{out}");
        let out = structure(&xml, &["S!3"], &[]).unwrap();
        assert!(out.contains(r#"table:number-rows-repeated="4""#), "{out}");
    }

    #[test]
    fn deleted_columns_shorten_cells_and_column_definitions() {
        let xml = doc(
            r#"<table:table table:name="S"><table:table-column table:number-columns-repeated="4"/><table:table-column/><table:table-row><table:table-cell table:number-columns-repeated="3"><text:p>a</text:p></table:table-cell><table:table-cell><text:p>b</text:p></table:table-cell></table:table-row></table:table>"#,
        );
        // Columns B-C (1-2) go.
        let out = structure(&xml, &[], &["S!B:C"]).unwrap();
        assert!(
            out.contains(r#"<table:table-column table:number-columns-repeated="2"/>"#),
            "{out}"
        );
        assert!(
            out.contains(r#"<table:table-cell table:number-columns-repeated="1""#)
                || out.contains("<table:table-cell><text:p>a</text:p>"),
            "{out}"
        );
        assert!(out.contains("<text:p>b</text:p>"), "{out}");
    }

    #[test]
    fn a_merged_area_shrinks_with_the_rows_it_loses() {
        let xml = doc(
            r#"<table:table table:name="S"><table:table-row><table:table-cell table:number-rows-spanned="3" table:number-columns-spanned="2"><text:p>m</text:p></table:table-cell><table:covered-table-cell/></table:table-row><table:table-row><table:covered-table-cell/><table:covered-table-cell/></table:table-row><table:table-row><table:covered-table-cell/><table:covered-table-cell/></table:table-row></table:table>"#,
        );
        let out = structure(&xml, &["S!3"], &[]).unwrap();
        assert!(out.contains(r#"table:number-rows-spanned="2""#), "{out}");
        assert!(out.contains(r#"table:number-columns-spanned="2""#), "{out}");
        // Down to one row, the span goes.
        let out = structure(&xml, &["S!2:3"], &[]).unwrap();
        assert!(!out.contains("number-rows-spanned"), "{out}");
        // The first cell of an area that stays cannot go.
        let e = structure(&xml, &["S!1"], &[]).unwrap_err();
        assert!(format!("{e:#}").contains("merged area"), "{e:#}");
    }

    #[test]
    fn a_formula_that_changes_loses_its_result_and_the_names_follow() {
        let xml = doc(
            r#"<table:table table:name="S"><table:table-row><table:table-cell office:value-type="float" office:value="5" table:formula="of:=[.A3]*2"><text:p>10</text:p><office:annotation><text:p>note</text:p></office:annotation></table:table-cell></table:table-row><table:table-row><table:table-cell/></table:table-row><table:table-row><table:table-cell office:value-type="float" office:value="1" table:formula="of:=[.B1]"><text:p>1</text:p></table:table-cell></table:table-row></table:table><table:named-expressions><table:named-range table:name="N" table:base-cell-address="$S.$A$1" table:cell-range-address="$S.$A$3:.$A$5"/><table:named-range table:name="Gone" table:base-cell-address="$S.$A$1" table:cell-range-address="$S.$A$2"/><table:named-expression table:name="E" table:base-cell-address="$S.$A$1" table:expression="[$S.A3]+[$S.A2]"/></table:named-expressions><table:database-ranges><table:database-range table:name="D" table:target-range-address="S.A2:S.B2"/><table:database-range table:name="Kept" table:target-range-address="S.A3:S.B4"/></table:database-ranges>"#,
        );
        let out = structure(&xml, &["S!2"], &[]).unwrap();
        // `[.A3]` moves to `[.A2]` and has no value; its note stays; the formula of the last row
        // is `[.B1]`, which does not change, and keeps its value.
        assert!(out.contains(r#"table:formula="of:=[.A2]*2""#), "{out}");
        assert!(
            !out.contains(r#"office:value="5""#) && !out.contains("<text:p>10</text:p>"),
            "{out}"
        );
        assert!(out.contains("<text:p>note</text:p>"), "{out}");
        assert!(
            out.contains(r#"office:value="1" table:formula="of:=[.B1]""#),
            "{out}"
        );
        // The names: a range follows, one that is gone is `#REF!`, an expression is adjusted.
        assert!(
            out.contains(r#"table:cell-range-address="$S.$A$2:.$A$4""#),
            "{out}"
        );
        assert!(out.contains(r##"table:name="Gone" table:base-cell-address="$S.$A$1" table:cell-range-address="#REF!""##), "{out}");
        assert!(
            out.contains(r#"table:expression="[$S.A2]+[#REF!]""#),
            "{out}"
        );
        // A database range with no cell left goes, the other moves.
        assert!(!out.contains(r#"table:name="D""#), "{out}");
        assert!(
            out.contains(r#"table:target-range-address="S.A2:S.B3""#),
            "{out}"
        );
    }

    #[test]
    fn a_sheet_removed_turns_its_references_into_ref_errors() {
        let xml = doc(
            r#"<table:table table:name="T"><table:table-row><table:table-cell/></table:table-row></table:table><table:table table:name="F"><table:table-row><table:table-cell table:formula="of:=SUM([$T.A1:.A3];[.B1])"/></table:table-row></table:table>"#,
        );
        let out = modify_xml(
            &xml,
            &WorkbookModifications {
                remove_sheets: Some(HashSet::from(["T".to_string()])),
                ..Default::default()
            },
        )
        .unwrap();
        assert!(
            out.contains(r#"table:formula="of:=SUM([#REF!];[.B1])""#),
            "{out}"
        );
    }

    #[test]
    fn recalculation_on_open_drops_every_cached_result() {
        let xml = doc(
            r#"<table:table table:name="S"><table:table-row><table:table-cell office:value-type="float" office:value="5" table:formula="of:=[.B1]"><text:p>5</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell></table:table-row></table:table>"#,
        );
        let out = modify_xml(
            &xml,
            &WorkbookModifications {
                recalculate_on_open: true,
                ..Default::default()
            },
        )
        .unwrap();
        // The formula has no result; the constant keeps its value.
        assert!(!out.contains(r#"office:value="5""#), "{out}");
        assert!(
            out.contains(r#"table:formula="of:=[.B1]""#) && out.contains(r#"office:value="7""#),
            "{out}"
        );
    }

    #[test]
    fn what_cannot_be_adjusted_stops_the_command() {
        let xml = doc(
            r#"<table:table table:name="S"><table:table-row><table:table-cell table:formula="of:=MUNIT(2)" table:number-matrix-columns-spanned="2" table:number-matrix-rows-spanned="2"/><table:covered-table-cell/></table:table-row><table:table-row><table:covered-table-cell/><table:covered-table-cell/></table:table-row></table:table>"#,
        );
        let e = structure(&xml, &["S!2"], &[]).unwrap_err();
        assert!(format!("{e:#}").contains("array formula"), "{e:#}");
        // A sheet that is not in the document.
        let e = structure(&xml, &["Nope!1"], &[]).unwrap_err();
        assert!(format!("{e:#}").contains("not found"), "{e:#}");
        // A formula in another syntax is not read.
        let other = doc(
            r#"<table:table table:name="S"><table:table-row><table:table-cell table:formula="oooc:=[.A2]"/></table:table-row><table:table-row/></table:table>"#,
        );
        assert!(structure(&other, &["S!2"], &[]).is_err());
    }

    #[test]
    fn pivot_tables_and_filters_stop_a_deletion_of_rows_or_columns() {
        let table = r#"<table:table table:name="S"><table:table-row/></table:table>"#;
        let pivot = doc(&format!(
            r#"{table}<table:data-pilot-tables><table:data-pilot-table table:name="P"/></table:data-pilot-tables>"#
        ));
        let e = structure(&pivot, &["S!1"], &[]).unwrap_err();
        assert!(format!("{e:#}").contains("pivot tables"), "{e:#}");
        let filter = doc(&format!(
            r#"{table}<table:database-ranges><table:database-range table:name="D" table:target-range-address="S.A1:S.B2"><table:filter/></table:database-range></table:database-ranges>"#
        ));
        assert!(structure(&filter, &[], &["S!A"]).is_err());
        // Without a deletion of rows or columns they are not in the way.
        let removed = modify_xml(
            &pivot,
            &WorkbookModifications {
                recalculate_on_open: true,
                ..Default::default()
            },
        );
        assert!(removed.is_ok());
    }

    #[test]
    fn a_content_validation_reads_the_sheet_of_its_base_cell() {
        let xml = doc(
            r#"<table:content-validations><table:content-validation table:name="v" table:condition="of:cell-content-is-between([.A3];5)" table:base-cell-address="S.A1"/></table:content-validations><table:table table:name="S"><table:table-row/><table:table-row/><table:table-row/></table:table>"#,
        );
        let out = structure(&xml, &["S!2"], &[]).unwrap();
        assert!(out.contains("of:cell-content-is-between([.A2];5)"), "{out}");
    }

    /// The shape that is anchored to a cell that goes stays where it was: it belongs to the cell
    /// that has the index of the one it left (what LibreOffice 26.8 gave, see
    /// `sheetrs-core/scripts/compare_structure_with_libreoffice.py`).
    #[test]
    fn a_shape_of_a_deleted_cell_goes_to_the_cell_that_takes_its_place() {
        let cell =
            |text: &str| format!("<table:table-cell><text:p>{text}</text:p></table:table-cell>");
        let shape = r#"<table:table-cell><draw:rect/></table:table-cell>"#;
        let row = |cells: &str| format!("<table:table-row>{cells}</table:table-row>");
        let rows = [
            row(&cell("a")),
            row(shape),
            row(&cell("c")),
            row(&cell("d")),
        ]
        .concat();
        let xml = doc(&format!(
            r#"<table:table table:name="S">{rows}</table:table>"#
        ));
        // Row 2 goes: its shape is in the cell of row 2, which is "c" now.
        let out = structure(&xml, &["S!2"], &[]).unwrap();
        assert!(
            out.contains("<text:p>a</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell><draw:rect/><text:p>c</text:p>"),
            "{out}"
        );
        // A shape in a cell that stays moves with it.
        let out = structure(&xml, &["S!1"], &[]).unwrap();
        assert_eq!(out.matches("<draw:rect/>").count(), 1, "{out}");
        // Rows 3-4 go: the row of the shape is the last one.
        let out = structure(&xml, &["S!3:4"], &[]).unwrap();
        assert!(!out.contains("<text:p>c</text:p>"), "{out}");
        // Row 2 and the rows after it go: the shape needs a row that no input row gives.
        let out = structure(&xml, &["S!2:4"], &[]).unwrap();
        assert!(!out.contains("<text:p>c</text:p>"), "{out}");
        assert!(
            out.ends_with(&format!("{}</table:table>{DOC_TAIL}", row(shape))),
            "{out}"
        );
    }

    #[test]
    fn the_shapes_past_the_last_row_keep_their_rows_with_empty_rows_between() {
        let row = |inner: &str| format!("<table:table-row>{inner}</table:table-row>");
        let cell = |inner: &str| format!("<table:table-cell>{inner}</table:table-cell>");
        let rows = [
            row(&cell("a")),
            row(&cell(r#"<draw:rect draw:name="one"/>"#)),
            row(&cell("c")),
            row(&cell(r#"<draw:rect draw:name="two"/>"#)),
            row(&cell("e")),
        ]
        .concat();
        let xml = doc(&format!(
            r#"<table:table table:name="S">{rows}</table:table>"#
        ));
        // Row 2 and rows 4-5 go: "one" is in the cell that "c" takes, "two" is anchored to row 4,
        // two rows after the last one that stays.
        let out = structure(&xml, &["S!2", "S!4:5"], &[]).unwrap();
        assert!(
            out.contains(r#"<table:table-cell><draw:rect draw:name="one"/>c</table:table-cell>"#),
            "{out}"
        );
        assert!(
            out.ends_with(&format!(
                "{}{}</table:table>{DOC_TAIL}",
                row("<table:table-cell/>"),
                row(&cell(r#"<draw:rect draw:name="two"/>"#))
            )),
            "{out}"
        );
    }

    #[test]
    fn a_shape_in_a_deleted_column_of_repeated_rows_is_kept_once() {
        let xml = doc(
            r#"<table:table table:name="S"><table:table-row table:number-rows-repeated="3"><table:table-cell/><table:table-cell><draw:rect/></table:table-cell><table:table-cell/></table:table-row></table:table>"#,
        );
        let out = structure(&xml, &[], &["S!B"]).unwrap();
        assert_eq!(out.matches("<draw:rect/>").count(), 1, "{out}");
    }

    #[test]
    fn a_shape_in_a_run_of_repeated_rows_splits_it() {
        let xml = doc(
            r#"<table:table table:name="S"><table:table-row><table:table-cell><draw:rect/></table:table-cell></table:table-row><table:table-row table:number-rows-repeated="5"><table:table-cell table:number-columns-repeated="4"/></table:table-row></table:table>"#,
        );
        // The first row goes: the shape is in the first cell of the first row of the run.
        let out = structure(&xml, &["S!1"], &[]).unwrap();
        assert!(
            out.contains(r#"<table:table-row><table:table-cell><draw:rect/></table:table-cell><table:table-cell table:number-columns-repeated="3"/></table:table-row><table:table-row table:number-rows-repeated="4"><table:table-cell table:number-columns-repeated="4"/></table:table-row>"#),
            "{out}"
        );
    }

    #[test]
    fn a_shape_of_a_deleted_column_goes_to_the_cell_that_takes_its_place() {
        let xml = doc(
            r#"<table:table table:name="S"><table:table-row><table:table-cell/><table:table-cell><draw:rect/></table:table-cell><table:table-cell><text:p>c</text:p></table:table-cell></table:table-row></table:table>"#,
        );
        let out = structure(&xml, &[], &["S!B"]).unwrap();
        assert!(
            out.contains("<table:table-cell/><table:table-cell><draw:rect/><text:p>c</text:p></table:table-cell>"),
            "{out}"
        );
    }

    #[test]
    fn a_shape_whose_cell_goes_with_the_rows_before_it_stops_the_command() {
        let xml = doc(
            r#"<table:table table:name="S"><table:table-row/><table:table-row/><table:table-row><table:table-cell><draw:rect/></table:table-cell></table:table-row><table:table-row/></table:table>"#,
        );
        // Rows 2-3 go, and the shape is in row 3: LibreOffice places it by sizes the file lacks.
        let e = structure(&xml, &["S!2:3"], &[]).unwrap_err();
        assert!(format!("{e:#}").contains("height of the rows"), "{e:#}");
        // The band that starts at the shape's row, or below it, is fine.
        assert!(structure(&xml, &["S!3:4"], &[]).is_ok());
        assert!(structure(&xml, &["S!2"], &[]).is_ok());
    }

    #[test]
    fn the_end_cell_of_a_shape_moves_with_its_cell() {
        let xml = doc(
            r#"<table:table table:name="S"><table:table-row/><table:table-row><table:table-cell><draw:rect table:end-cell-address="S.C4" table:end-x="1cm"/></table:table-cell></table:table-row></table:table>"#,
        );
        // The shape's cell is A2 and it ends in C4: with row 1 gone it is in A1 and ends in C3.
        let out = structure(&xml, &["S!1"], &[]).unwrap();
        assert!(out.contains(r#"table:end-cell-address="S.C3""#), "{out}");
        assert!(out.contains(r#"table:end-x="1cm""#), "{out}");
    }
}

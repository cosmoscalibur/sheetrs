//! ODS writer functionality for removing sheets and named ranges, and for
//! converting formulas to values

use super::{
    ValueScope, ValueTarget, WorkbookModifications, attribute, read_element_body, write_element,
};
use anyhow::{Context, Result, bail};
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashSet;
use std::io::{BufRead, BufReader, BufWriter, Read, Seek, Write};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

/// Part holding the sheets and named ranges of an ODS document.
const CONTENT_PART: &str = "content.xml";

const ROW: &str = "table:table-row";
const ROWS_REPEATED: &str = "table:number-rows-repeated";
const COLUMNS_REPEATED: &str = "table:number-columns-repeated";
const FORMULA: &str = "table:formula";
const MATRIX_COLUMNS: &str = "table:number-matrix-columns-spanned";
const MATRIX_ROWS: &str = "table:number-matrix-rows-spanned";

/// Modify an ODS file by applying specified modifications
pub fn modify_workbook_ods<R: Read + Seek, W: Write + Seek>(
    input: R,
    output: W,
    modifications: &WorkbookModifications,
) -> Result<()> {
    let mut archive = ZipArchive::new(input)?;
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
        )
        .with_context(|| format!("Failed to modify {CONTENT_PART}"))?;
    }

    zip_writer.finish()?;
    Ok(())
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

/// Apply the modifications to `content.xml`: drop removed sheets and named
/// ranges, and convert formulas to values in the sheets of `convert_to_values`.
fn modify_content_xml<R: BufRead, W: Write>(
    input: R,
    output: W,
    modifications: &WorkbookModifications,
) -> Result<()> {
    let removes_sheet = |name: &str| {
        modifications
            .remove_sheets
            .as_ref()
            .is_some_and(|sheets| sheets.contains(name))
    };
    let removes_range = |name: &str| {
        modifications
            .remove_named_ranges
            .as_ref()
            .is_some_and(|ranges| ranges.contains(name))
    };

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
        match reader.read_event_into(&mut buf)? {
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
                    sheet = name;
                    row = 0;
                }
                writer.write_event(Event::Start(e))?;
            }
            Event::End(e) if e.name().as_ref() == "table:table" => {
                table_depth = table_depth.saturating_sub(1);
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
                if scopes.is_empty() {
                    writer.write_event(Event::Start(e))?;
                } else {
                    let body = read_element_body(&mut reader, ROW)?;
                    write_rows(&mut writer, &e, &body, row, repeated, &scopes)?;
                }
                row = row.saturating_add(repeated);
            }
            Event::Empty(e) if e.name().as_ref() == ROW && table_depth == 1 => {
                row = row.saturating_add(repeat_count(&e, ROWS_REPEATED)?);
                writer.write_event(Event::Empty(e))?;
            }
            Event::Start(e) if is_named_range(e.name().as_ref()) => {
                if removes_range(&attribute(&e, "table:name")?.unwrap_or_default()) {
                    let end = e.to_end().into_owned();
                    reader.read_to_end_into(end.name(), &mut skip_buf)?;
                } else {
                    writer.write_event(Event::Start(e))?;
                }
            }
            Event::Empty(e) if is_named_range(e.name().as_ref()) => {
                if !removes_range(&attribute(&e, "table:name")?.unwrap_or_default()) {
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
    for scope in &modifications.convert_to_values {
        if !kept_sheets.contains(&scope.sheet) && !removes_sheet(&scope.sheet) {
            bail!("Sheet '{}' not found in {CONTENT_PART}", scope.sheet);
        }
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

/// Write a `table:table-row` covering `repeated` rows from `first_row`,
/// converting the formulas of its cells that fall inside `scopes`.
///
/// Repeated rows and cells are split where a scope starts or ends, so each
/// written element lies wholly inside or outside every scope.
fn write_rows<W: Write>(
    writer: &mut Writer<W>,
    start: &BytesStart,
    body: &[Event],
    first_row: u32,
    repeated: u32,
    scopes: &[&ValueScope],
) -> Result<()> {
    let items = row_items(body);
    let mut has_formula = false;
    for item in &items {
        if let RowItem::Cell { start, .. } = item
            && start.try_get_attribute(FORMULA)?.is_some()
        {
            has_formula = true;
            break;
        }
    }
    if !has_formula {
        writer.write_event(Event::Start(start.borrow()))?;
        for event in body {
            writer.write_event(event.borrow())?;
        }
        writer.write_event(Event::End(start.to_end()))?;
        return Ok(());
    }

    let row_bands = scopes.iter().filter_map(|scope| match scope.target {
        ValueTarget::Rows(first, last) => Some((first, last)),
        _ => None,
    });
    for (row, count) in split_repeats(first_row, repeated, row_bands) {
        writer.write_event(Event::Start(rebuild_start(
            start,
            ROWS_REPEATED,
            count,
            |_| false,
        )?))?;
        let mut col = 0u32;
        for item in &items {
            match item {
                RowItem::Cell { start, content } => {
                    let repeated = repeat_count(start, COLUMNS_REPEATED)?;
                    write_cells(writer, start, *content, row, col, repeated, scopes)?;
                    col = col.saturating_add(repeated);
                }
                RowItem::Other(event) => writer.write_event(event.borrow())?,
            }
        }
        writer.write_event(Event::End(start.to_end()))?;
    }
    Ok(())
}

/// Write a cell covering `repeated` columns from `first_col` of `row`.
fn write_cells<W: Write>(
    writer: &mut Writer<W>,
    start: &BytesStart,
    content: Option<&[Event]>,
    row: u32,
    first_col: u32,
    repeated: u32,
    scopes: &[&ValueScope],
) -> Result<()> {
    let has_formula = attribute(start, FORMULA)?.is_some();
    // LibreOffice stores an empty text result as a paragraph without a value
    // type. A formula with neither has no cached value to convert to.
    let value_type = attribute(start, "office:value-type")?;
    let has_paragraph = content.is_some_and(|events| {
        events.iter().any(|event| {
            matches!(event, Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "text:p")
        })
    });
    if !has_formula || (value_type.is_none() && !has_paragraph) {
        return write_element(writer, start.clone(), content);
    }

    let is_error = attribute(start, "calcext:value-type")?.as_deref() == Some("error");
    let matrix_cols = repeat_count(start, MATRIX_COLUMNS)?;
    let matrix_rows = repeat_count(start, MATRIX_ROWS)?;
    let column_bands = scopes.iter().filter_map(|scope| match scope.target {
        ValueTarget::Columns(first, last) => Some((first, last)),
        _ => None,
    });
    for (col, count) in split_repeats(first_col, repeated, column_bands) {
        // Array formulas convert only when one scope covers their whole range,
        // since their other cells hold no formula of their own.
        let last = (row + matrix_rows - 1, col + matrix_cols - 1);
        let converts = scopes
            .iter()
            .any(|scope| scope.contains_range((row, col), last));
        let cell = if converts {
            let mut cell = rebuild_start(start, COLUMNS_REPEATED, count, |key| {
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
            rebuild_start(start, COLUMNS_REPEATED, count, |_| false)?
        };
        write_element(writer, cell, content)?;
    }
    Ok(())
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
        modify_content_xml(xml.as_bytes(), &mut output, modifications)?;
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
}

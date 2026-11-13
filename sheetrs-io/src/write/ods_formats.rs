//! The cell styles of an ODS that a replacement needs when it gives a cell a number format.
//!
//! The automatic styles come before the cells in `content.xml`, and which style a cell has (its
//! own, or the default of its row or column) is only known on reading them. So a first pass reads
//! the styles and the cells to be formatted, and makes each cell style that is needed: the style
//! the cell has, with the data style of the format. The writer then adds them at the end of the
//! automatic styles and the cells take them.

use super::error::Result;
use super::formats::{ods_data_style, ods_value_type};
use super::{CellEdit, SheetEdits, attribute};
use crate::EntryFormat;
use crate::reader::parser_utils::attr_value;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::fmt::Write as _;
use std::io::BufRead;

/// The cell styles and the edits to write.
pub(crate) struct Resolved {
    /// The styles to add to `office:automatic-styles`.
    pub styles: String,
    /// The edits of the sheets with the style of each formatted cell.
    pub replacements: HashMap<String, SheetEdits>,
}

/// An automatic style of the family `table-cell`.
struct CellStyle {
    attributes: Vec<(String, String)>,
    inner: String,
}

fn repeat(element: &BytesStart, key: &str) -> u32 {
    attribute(element, key)
        .ok()
        .flatten()
        .and_then(|n| n.parse::<u32>().ok())
        .unwrap_or(1)
        .max(1)
}

/// Whether some replacement gives a format.
pub(crate) fn wanted(replacements: &HashMap<String, SheetEdits>) -> bool {
    replacements.values().any(|edits| {
        edits
            .values()
            .any(|e| matches!(e, CellEdit::Formatted { .. }))
    })
}

/// Reads `content.xml` and makes the styles of the formatted cells of `replacements`.
pub(crate) fn resolve<R: BufRead>(
    input: R,
    replacements: &HashMap<String, SheetEdits>,
) -> Result<Resolved> {
    let mut reader = Reader::from_reader(input);
    let mut cell_styles: HashMap<String, CellStyle> = HashMap::new();
    // The style a formatted cell has: `(row, column)` of a sheet.
    let mut parents: HashMap<(String, u32, u32), Option<String>> = HashMap::new();
    let mut in_automatic = false;
    let mut capturing: Option<(String, CellStyle, u32)> = None;
    let mut inner = Writer::new(Vec::new());
    let mut depth = 0u32;
    let mut sheet = String::new();
    let mut columns: Vec<(u32, u32, Option<String>)> = Vec::new();
    let mut next_column = 0u32;
    let mut row = 0u32;
    let mut row_repeat = 1u32;
    let mut row_default: Option<String> = None;
    let mut col = 0u32;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        let event = reader.read_event_into(&mut buf)?;
        if matches!(event, Event::Eof) {
            break;
        }
        // The automatic styles of the cells.
        match &event {
            Event::Start(e) if e.name().as_ref() == "office:automatic-styles" => {
                in_automatic = true;
            }
            Event::End(e) if e.name().as_ref() == "office:automatic-styles" => {
                in_automatic = false;
            }
            Event::Start(e) | Event::Empty(e)
                if in_automatic
                    && capturing.is_none()
                    && e.name().as_ref() == "style:style"
                    && attribute(e, "style:family")?.as_deref() == Some("table-cell") =>
            {
                let name = attribute(e, "style:name")?.unwrap_or_default();
                let attributes = e
                    .attributes()
                    .flatten()
                    .filter_map(|a| {
                        let value = attr_value(&a).ok()?.into_owned();
                        Some((a.key.as_ref().to_string(), value))
                    })
                    .collect();
                let style = CellStyle {
                    attributes,
                    inner: String::new(),
                };
                if matches!(event, Event::Empty(_)) {
                    cell_styles.insert(name, style);
                } else {
                    inner = Writer::new(Vec::new());
                    capturing = Some((name, style, 1));
                }
                continue;
            }
            _ => {}
        }
        if let Some((name, style, nested)) = capturing.as_mut() {
            match &event {
                Event::Start(_) => *nested += 1,
                Event::End(_) => *nested -= 1,
                _ => {}
            }
            if *nested == 0 {
                let name = std::mem::take(name);
                let mut style = CellStyle {
                    attributes: std::mem::take(&mut style.attributes),
                    inner: String::new(),
                };
                style.inner =
                    String::from_utf8_lossy(&std::mem::take(inner.get_mut())).into_owned();
                cell_styles.insert(name, style);
                capturing = None;
            } else {
                inner.write_event(event.borrow())?;
            }
            continue;
        }
        // The cells.
        match &event {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "table:table" => {
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                    if depth == 1 {
                        sheet = attribute(e, "table:name")?.unwrap_or_default();
                        columns.clear();
                        next_column = 0;
                        row = 0;
                    }
                }
            }
            Event::End(e) if e.name().as_ref() == "table:table" => depth = depth.saturating_sub(1),
            _ if depth != 1 => {}
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "table:table-column" => {
                let count = repeat(e, "table:number-columns-repeated");
                columns.push((
                    next_column,
                    next_column.saturating_add(count),
                    attribute(e, "table:default-cell-style-name")?,
                ));
                next_column = next_column.saturating_add(count);
            }
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == "table:table-row" => {
                row_repeat = repeat(e, "table:number-rows-repeated");
                row_default = attribute(e, "table:default-cell-style-name")?;
                col = 0;
                if matches!(event, Event::Empty(_)) {
                    row += row_repeat;
                }
            }
            Event::End(e) if e.name().as_ref() == "table:table-row" => row += row_repeat,
            Event::Start(e) | Event::Empty(e)
                if matches!(
                    e.name().as_ref(),
                    "table:table-cell" | "table:covered-table-cell"
                ) =>
            {
                let count = repeat(e, "table:number-columns-repeated");
                if let Some(edits) = replacements.get(&sheet) {
                    let own = attribute(e, "table:style-name")?;
                    for (&(r, c), edit) in edits.range((row, col)..) {
                        if r >= row.saturating_add(row_repeat) {
                            break;
                        }
                        if c < col || c >= col.saturating_add(count) {
                            continue;
                        }
                        if matches!(edit, CellEdit::Formatted { .. }) {
                            let column_default = columns
                                .iter()
                                .find(|(from, to, _)| (*from..*to).contains(&c))
                                .and_then(|(_, _, style)| style.clone());
                            parents.insert(
                                (sheet.clone(), r, c),
                                own.clone()
                                    .or_else(|| row_default.clone())
                                    .or(column_default),
                            );
                        }
                    }
                }
                col = col.saturating_add(count);
            }
            _ => {}
        }
    }
    // The styles of what was asked, each once.
    let mut data_styles: HashMap<EntryFormat, String> = HashMap::new();
    let mut made: HashMap<(Option<String>, EntryFormat), String> = HashMap::new();
    let mut counter = 0usize;
    let mut styles = String::new();
    let mut resolved: HashMap<String, SheetEdits> = HashMap::new();
    for (name, edits) in replacements {
        let mut out = SheetEdits::new();
        for (&(r, c), edit) in edits {
            let CellEdit::Formatted {
                edit: inner,
                format,
                ods_value,
                shown,
                ..
            } = edit
            else {
                out.insert((r, c), edit.clone());
                continue;
            };
            let Some(parent) = parents.get(&(name.clone(), r, c)) else {
                out.insert((r, c), edit.clone());
                continue;
            };
            let format = *format;
            let data = data_styles
                .entry(format)
                .or_insert_with(|| {
                    let data = format!("sheetrs-data-{}", data_styles_len(format));
                    styles.push_str(&ods_data_style(format, &data));
                    data
                })
                .clone();
            let style = made
                .entry((parent.clone(), format))
                .or_insert_with(|| {
                    counter += 1;
                    let style = format!("sheetrs-cell-{counter}");
                    styles.push_str(&cell_style(&style, parent.as_deref(), &data, &cell_styles));
                    style
                })
                .clone();
            out.insert(
                (r, c),
                CellEdit::Formatted {
                    edit: inner.clone(),
                    format,
                    ods_value: ods_value.clone(),
                    shown: shown.clone(),
                    style: Some(style),
                },
            );
        }
        resolved.insert(name.clone(), out);
    }
    Ok(Resolved {
        styles,
        replacements: resolved,
    })
}

/// A number that tells the formats apart in a name.
fn data_styles_len(format: EntryFormat) -> usize {
    format as usize
}

/// A cell style that is `parent` (an automatic style, whose properties it copies, or a common
/// one, which it inherits) with the data style `data`.
fn cell_style(
    name: &str,
    parent: Option<&str>,
    data: &str,
    automatic: &HashMap<String, CellStyle>,
) -> String {
    let escape = |text: &str| {
        text.replace('&', "&amp;")
            .replace('"', "&quot;")
            .replace('<', "&lt;")
    };
    match parent.and_then(|p| automatic.get(p)) {
        Some(style) => {
            let mut open = format!(r#"<style:style style:name="{name}""#);
            for (key, value) in &style.attributes {
                if key != "style:name" && key != "style:data-style-name" {
                    let _ = write!(open, r#" {key}="{}""#, escape(value));
                }
            }
            let _ = write!(open, r#" style:data-style-name="{data}""#);
            if style.inner.is_empty() {
                format!("{open}/>")
            } else {
                format!("{open}>{}</style:style>", style.inner)
            }
        }
        None => format!(
            r#"<style:style style:name="{name}" style:family="table-cell" style:parent-style-name="{}" style:data-style-name="{data}"/>"#,
            escape(parent.unwrap_or("Default"))
        ),
    }
}

/// The attributes `office:value-type` and the value one, for a number with the format.
pub(crate) fn value_attributes(
    format: EntryFormat,
    x: f64,
    ods_value: Option<&str>,
) -> Vec<(&'static str, String)> {
    let kind = ods_value_type(format);
    let mut out = vec![("office:value-type", kind.to_string())];
    match (kind, ods_value) {
        ("date", Some(v)) => out.push(("office:date-value", v.to_string())),
        ("time", Some(v)) => out.push(("office:time-value", v.to_string())),
        ("currency", _) => {
            out.push(("office:currency", "USD".to_string()));
            out.push(("office:value", x.to_string()));
        }
        _ => out.push(("office:value", x.to_string())),
    }
    out.push(("calcext:value-type", kind.to_string()));
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    const CONTENT: &str = r#"<office:document-content><office:automatic-styles><style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold"/></style:style><style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N2"/></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="S"><table:table-column table:number-columns-repeated="2" table:default-cell-style-name="ce1"/><table:table-row><table:table-cell table:number-columns-repeated="2"/></table:table-row><table:table-row table:default-cell-style-name="ce2"><table:table-cell/><table:table-cell table:style-name="ce1"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;

    fn formatted(format: EntryFormat) -> CellEdit {
        CellEdit::Formatted {
            edit: Box::new(CellEdit::Number(1.0)),
            format,
            ods_value: None,
            shown: None,
            style: None,
        }
    }

    #[test]
    fn a_cell_takes_the_style_of_its_own_row_or_column_with_the_new_format() {
        let mut edits = SheetEdits::new();
        // No style of its own: the one of its column, a bold style that is copied.
        edits.insert((0, 1), formatted(EntryFormat::Percent));
        // The default of its row, a style that has a number format that is replaced.
        edits.insert((1, 0), formatted(EntryFormat::Percent));
        // Its own style.
        edits.insert((1, 1), formatted(EntryFormat::Date));
        let replacements = HashMap::from([("S".to_string(), edits)]);
        let resolved = resolve(CONTENT.as_bytes(), &replacements).unwrap();
        let style_of = |r, c| match &resolved.replacements["S"][&(r, c)] {
            CellEdit::Formatted { style, .. } => style.clone().unwrap(),
            _ => panic!("not formatted"),
        };
        assert_eq!(style_of(0, 1), "sheetrs-cell-1");
        assert_ne!(style_of(1, 0), style_of(0, 1));
        let styles = &resolved.styles;
        assert!(styles.contains(
            r#"<style:style style:name="sheetrs-cell-1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="sheetrs-data-0"><style:text-properties fo:font-weight="bold"/></style:style>"#
        ), "{styles}");
        assert!(styles.contains("<number:percentage-style"));
        assert!(styles.contains("<number:date-style"));
        // The style `ce2` had its own data style, which is not copied.
        assert!(!styles.contains(r#"style:data-style-name="N2""#));
    }
}

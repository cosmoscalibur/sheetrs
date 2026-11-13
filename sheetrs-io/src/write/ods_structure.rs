//! References of an ODS document under a structural plan: the formulas of the cells
//! (`of:=SUM([$T.A2:.A3])`), the named expressions and the attributes that hold ranges.
//!
//! The semantics are those of `sheetrs_formula::edit`, which reads Excel's A1 syntax; this module
//! reads and writes the notation of ODF (`[$Sheet.A1:.B2]`) around it.

use super::attribute;
use super::error::{Error, Result};
use super::shapes::{self, Kind, Sizes, Span};
use super::structure::three_d_unsupported;
use super::xlsx_structure::{SheetExtents, SheetStructure};
use quick_xml::events::{BytesStart, Event};
use sheetrs_formula::addr::{col_index, col_name, needs_quotes};
use sheetrs_formula::ast::{CellAddr, ColAddr, Expr, RefArea, Reference, RowAddr};
use sheetrs_formula::edit::apply;

/// A reference as an ODS formula writes it, inside `[` `]`.
struct OdsRef {
    /// The sheet is marked absolute (`$Sheet.`).
    sheet_abs: bool,
    /// The second end of a range repeats the sheet (`S.A1:S.B2`), as an attribute writes it,
    /// instead of leaving it out (`$S.A1:.B2`), as a formula does.
    repeats_sheet: bool,
    /// The last sheet of a 3D reference (`$A.A1:$C.B2`), when it is not the first. Such a
    /// reference is read to be refused.
    last_sheet: Option<String>,
    reference: Reference,
}

/// The sheet qualifier at the start of `text`, up to its `.`: `(absolute, name, rest)`. No sheet
/// (`.A1`) gives `None`.
fn sheet_part(text: &str) -> Option<(Option<(bool, String)>, &str)> {
    let (abs, body) = match text.strip_prefix('$') {
        Some(rest) => (true, rest),
        None => (false, text),
    };
    if let Some(rest) = body.strip_prefix('\'') {
        let mut name = String::new();
        let mut chars = rest.char_indices();
        while let Some((i, c)) = chars.next() {
            if c != '\'' {
                name.push(c);
            } else if rest[i + 1..].starts_with('\'') {
                name.push('\'');
                chars.next();
            } else {
                let after = rest[i + 1..].strip_prefix('.')?;
                return Some((Some((abs, name)), after));
            }
        }
        return None;
    }
    let dot = body.find('.')?;
    if dot == 0 {
        return (!abs).then(|| (None, &body[1..]));
    }
    let name = &body[..dot];
    (!name.contains(['[', ']', ':', '\'', ' ', '!']))
        .then(|| (Some((abs, name.to_string())), &body[dot + 1..]))
}

/// One end of an area: a cell, a column or a row.
enum Corner {
    Cell(CellAddr),
    Col(ColAddr),
    Row(RowAddr),
}

fn corner(text: &str) -> Option<Corner> {
    // `$3`: a whole row marked absolute, where the leading `$` is not a column's.
    if let Some(digits) = text.strip_prefix('$')
        && !digits.is_empty()
        && digits.bytes().all(|b| b.is_ascii_digit())
    {
        let row = digits
            .parse::<u32>()
            .ok()
            .filter(|r| (1..=sheetrs_formula::addr::MAX_ROWS).contains(r))?;
        return Some(Corner::Row(RowAddr {
            row: row - 1,
            abs: true,
        }));
    }
    let (col_abs, rest) = match text.strip_prefix('$') {
        Some(r) => (true, r),
        None => (false, text),
    };
    let letters = rest.bytes().take_while(u8::is_ascii_alphabetic).count();
    let (col, rest) = rest.split_at(letters);
    let (row_abs, digits) = match rest.strip_prefix('$') {
        Some(r) => (true, r),
        None => (false, rest),
    };
    if !digits.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let row = (!digits.is_empty())
        .then(|| {
            digits
                .parse::<u32>()
                .ok()
                .filter(|r| (1..=sheetrs_formula::addr::MAX_ROWS).contains(r))
        })
        .flatten();
    match (col.is_empty(), digits.is_empty(), row) {
        (false, false, Some(row)) => Some(Corner::Cell(CellAddr {
            row: row - 1,
            col: col_index(col)?,
            row_abs,
            col_abs,
        })),
        (false, true, _) if !row_abs => Some(Corner::Col(ColAddr {
            col: col_index(col)?,
            abs: col_abs,
        })),
        (true, false, Some(row)) if !col_abs => Some(Corner::Row(RowAddr {
            row: row - 1,
            abs: row_abs,
        })),
        _ => None,
    }
}

fn parse_ref(content: &str) -> Option<OdsRef> {
    let (first_sheet, rest) = sheet_part(content)?;
    let mut repeats_sheet = false;
    let mut last_sheet = None;
    let (first, last) = match rest.split_once(':') {
        Some((a, b)) => {
            // The second end repeats the sheet or leaves it out (`.B2`); another sheet is a 3D
            // reference.
            let (second_sheet, b) = sheet_part(b)?;
            repeats_sheet = second_sheet.is_some();
            if let Some((_, name)) = second_sheet
                && first_sheet.as_ref().is_none_or(|(_, first)| *first != name)
            {
                first_sheet.as_ref()?;
                last_sheet = Some(name);
            }
            (a, Some(b))
        }
        None => (rest, None),
    };
    let area = match (corner(first)?, last.map(corner)) {
        (Corner::Cell(a), None) => RefArea::Cell(a),
        (Corner::Cell(a), Some(Some(Corner::Cell(b)))) => RefArea::Range(a, b),
        (Corner::Col(a), Some(Some(Corner::Col(b)))) => RefArea::Cols(a, b),
        (Corner::Row(a), Some(Some(Corner::Row(b)))) => RefArea::Rows(a, b),
        _ => return None,
    };
    let (sheet_abs, sheet) = match first_sheet {
        Some((abs, name)) => (abs, Some(name)),
        None => (false, None),
    };
    Some(OdsRef {
        sheet_abs,
        repeats_sheet,
        last_sheet,
        reference: Reference { sheet, area },
    })
}

fn quoted(name: &str) -> String {
    if needs_quotes(name) {
        format!("'{}'", name.replace('\'', "''"))
    } else {
        name.to_string()
    }
}

fn cell_text(c: &CellAddr) -> String {
    format!(
        "{}{}{}{}",
        if c.col_abs { "$" } else { "" },
        col_name(c.col),
        if c.row_abs { "$" } else { "" },
        c.row + 1
    )
}

fn print_ref(r: &OdsRef) -> String {
    let sheet = match &r.reference.sheet {
        Some(name) => format!("{}{}", if r.sheet_abs { "$" } else { "" }, quoted(name)),
        None => String::new(),
    };
    // The second end of a range: `.B2`, or `S.B2` when the text repeated the sheet.
    let second = if let Some(last) = &r.last_sheet {
        format!("{}{}.", if r.sheet_abs { "$" } else { "" }, quoted(last))
    } else if r.repeats_sheet {
        format!(
            "{}.",
            r.reference.sheet.as_deref().map_or(String::new(), quoted)
        )
    } else {
        ".".to_string()
    };
    let abs = |a: bool| if a { "$" } else { "" };
    let area = match r.reference.area {
        RefArea::Cell(c) => cell_text(&c),
        RefArea::Range(a, b) => format!("{}:{second}{}", cell_text(&a), cell_text(&b)),
        RefArea::Cols(a, b) => format!(
            "{}{}:{second}{}{}",
            abs(a.abs),
            col_name(a.col),
            abs(b.abs),
            col_name(b.col)
        ),
        RefArea::Rows(a, b) => format!(
            "{}{}:{second}{}{}",
            abs(a.abs),
            a.row + 1,
            abs(b.abs),
            b.row + 1
        ),
    };
    format!("{sheet}.{area}")
}

/// The ODS notation of the reference `content` (the text between `[` and `]`) once the plan is
/// applied: `Some(None)` when it is gone (`#REF!`), `Some(Some(text))` when it changes, `None`
/// when it stays or is not a reference this reads.
#[allow(clippy::option_option)] // three outcomes named above, told apart by every caller
fn adjusted(s: &SheetStructure, content: &str) -> Result<Option<Option<String>>> {
    adjusted_tracked(s, content, &mut true)
}

/// As [`adjusted`]; `only_moved` becomes `false` when the reference does not point to cells of the
/// same shape any more (it is gone, or its range shrank).
#[allow(clippy::option_option)]
fn adjusted_tracked(
    s: &SheetStructure,
    content: &str,
    only_moved: &mut bool,
) -> Result<Option<Option<String>>> {
    let Some(parsed) = parse_ref(content) else {
        return Ok(None);
    };
    if parsed.last_sheet.is_some() {
        return Err(three_d_unsupported(content));
    }
    let mut node = Expr::Ref(parsed.reference.clone());
    let mut changed = false;
    for edit in s.edits {
        changed |= apply(&mut node, edit, &s.sheet);
    }
    if !changed {
        return Ok(None);
    }
    if !matches!(&node, Expr::Ref(after)
        if after.sheet == parsed.reference.sheet
            && sheetrs_formula::edit::same_shape(&parsed.reference.area, &after.area))
    {
        *only_moved = false;
    }
    Ok(Some(match node {
        Expr::Ref(reference) => Some(print_ref(&OdsRef {
            sheet_abs: parsed.sheet_abs,
            repeats_sheet: parsed.repeats_sheet,
            last_sheet: None,
            reference,
        })),
        _ => None,
    }))
}

/// The text of the ODS formula (or expression) `text` once the plan is applied, `None` when it
/// does not change. References in brackets are adjusted; one that goes is `[#REF!]`.
pub(super) fn rewrite_expression(s: &SheetStructure, text: &str) -> Result<Option<String>> {
    Ok(rewrite_expression_tracked(s, text)?.map(|(text, _)| text))
}

/// As [`rewrite_expression`], and whether the plan only moved what the expression reads.
pub(super) fn rewrite_expression_tracked(
    s: &SheetStructure,
    text: &str,
) -> Result<Option<(String, bool)>> {
    let mut only_moved = true;
    let bytes = text.as_bytes();
    let mut out = String::with_capacity(text.len());
    let mut changed = false;
    let mut i = 0;
    let mut copied = 0;
    while i < bytes.len() {
        match bytes[i] {
            b'"' => {
                // Text between quotes, a doubled quote standing for one inside it.
                i += 1;
                while i < bytes.len() {
                    if bytes[i] == b'"' {
                        if bytes.get(i + 1) == Some(&b'"') {
                            i += 2;
                            continue;
                        }
                        break;
                    }
                    i += 1;
                }
                i += 1;
            }
            b'[' => {
                // A quoted sheet name can hold `]`.
                let mut j = i + 1;
                let mut in_quote = false;
                while j < bytes.len() && (in_quote || bytes[j] != b']') {
                    if bytes[j] == b'\'' {
                        in_quote = !in_quote;
                    }
                    j += 1;
                }
                if j >= bytes.len() {
                    break;
                }
                if let Some(new) = adjusted_tracked(s, &text[i + 1..j], &mut only_moved)? {
                    out.push_str(&text[copied..i]);
                    out.push('[');
                    out.push_str(new.as_deref().unwrap_or("#REF!"));
                    out.push(']');
                    copied = j + 1;
                    changed = true;
                }
                i = j + 1;
            }
            _ => i += 1,
        }
    }
    out.push_str(&text[copied..]);
    Ok(changed.then_some((out, only_moved)))
}

/// The formula of a `table:formula` attribute (`of:=…`) once the plan is applied, and whether the
/// plan only moved what the formula reads: the result it cached is then still its result.
pub(super) fn rewrite_cell_formula_tracked(
    s: &SheetStructure,
    formula: &str,
) -> Result<Option<(String, bool)>> {
    if !s.changes_layout() && !s.plan.may_reach_edited_sheet(formula) {
        return Ok(None);
    }
    let Some(body) = formula.strip_prefix("of:=") else {
        return Err(Error::malformed(format!(
            "the formula '{formula}' is not in the OpenFormula syntax (of:=)"
        )));
    };
    Ok(rewrite_expression_tracked(s, body)?.map(|(new, moved)| (format!("of:={new}"), moved)))
}

/// The ranges of a list separated by spaces; a space inside a quoted sheet name is not one.
fn split_ranges(value: &str) -> Vec<&str> {
    let mut parts = Vec::new();
    let (mut start, mut in_quote) = (0, false);
    for (i, c) in value.char_indices() {
        match c {
            '\'' => in_quote = !in_quote,
            ' ' if !in_quote => {
                parts.push(&value[start..i]);
                start = i + 1;
            }
            _ => {}
        }
    }
    parts.push(&value[start..]);
    parts
}

/// The sheet an address such as `$S.$A$1` is on, `None` when it names none.
pub(super) fn sheet_of_address(address: &str) -> Option<String> {
    sheet_part(address)?.0.map(|(_, name)| name)
}

/// A cell range attribute (`S.A1:.B2`, one or several separated by spaces) once the plan is
/// applied: the ranges that stay, `None` when none does.
pub(super) fn adjust_range_list(s: &SheetStructure, value: &str) -> Result<Option<String>> {
    let mut kept: Vec<String> = Vec::new();
    for range in split_ranges(value)
        .into_iter()
        .filter(|range| !range.is_empty())
    {
        match adjusted(s, range)? {
            None => kept.push(range.to_string()),
            Some(Some(new)) => kept.push(new),
            Some(None) => {}
        }
    }
    Ok((!kept.is_empty()).then(|| kept.join(" ")))
}

/// A length as ODF writes it (`0.8661in`, `2.2cm`, `12.8pt`, `5mm`, `10px`) in EMU.
fn length_emu(text: &str) -> Option<i64> {
    let text = text.trim();
    let split = text.find(|c: char| c.is_ascii_alphabetic())?;
    let (number, unit) = text.split_at(split);
    let number: f64 = number.trim().parse().ok()?;
    let per_unit = match unit {
        "in" => 914_400.0,
        "cm" => 360_000.0,
        "mm" => 36_000.0,
        "pt" => 12_700.0,
        "pc" => 152_400.0,
        "px" => 9_525.0,
        _ => return None,
    };
    Some((number * per_unit).round() as i64)
}

/// A length in EMU as ODF writes it, in inches.
fn emu_length(emu: i64) -> String {
    let text = format!("{:.5}", emu as f64 / 914_400.0);
    format!("{}in", text.trim_end_matches('0').trim_end_matches('.'))
}

/// `LibreOffice` keeps the sizes of rows and columns in twips (635 EMU).
fn in_twips(emu: i64) -> i64 {
    (emu as f64 / 635.0).round() as i64 * 635
}

/// The row and column styles that start a new page (`fo:break-before="page"`), each with the style
/// that is the same without the break: what the first row or column of a sheet gets when a deletion
/// leaves a break before it. `LibreOffice` keeps such a break in memory and reads none from a file.
#[derive(Default)]
pub(super) struct NoBreakStyles {
    /// The row styles with a break, to the name of their variant without it.
    pub rows: std::collections::HashMap<String, String>,
    /// The column styles with a break, to the name of their variant without it.
    pub cols: std::collections::HashMap<String, String>,
    /// The variants, as automatic styles to add to the document (those it does not have already).
    pub xml: String,
    /// The variants noted, with their XML, in the order of the document.
    variants: Vec<(String, String)>,
}

impl NoBreakStyles {
    /// Notes the style `name` of the family if its properties `e` break the page, and writes its
    /// variant.
    fn note(&mut self, name: &str, family: &str, e: &BytesStart) -> Result<()> {
        if attribute(e, "fo:break-before")?.as_deref() != Some("page") {
            return Ok(());
        }
        let variant = format!("{name}-nobreak");
        let mut properties = BytesStart::new(e.name().as_ref().to_string());
        for attr in e.attributes() {
            let attr = attr.map_err(quick_xml::Error::from)?;
            if attr.key.as_ref() == "fo:break-before" {
                properties.push_attribute(("fo:break-before", "auto"));
            } else {
                properties.push_attribute(attr);
            }
        }
        let mut properties_text = String::new();
        {
            let mut writer = quick_xml::Writer::new(Vec::new());
            writer
                .write_event(Event::Empty(properties))
                .map_err(quick_xml::Error::from)?;
            properties_text.push_str(&String::from_utf8_lossy(&writer.into_inner()));
        }
        self.variants.push((
            variant.clone(),
            format!(
                "<style:style style:name=\"{variant}\" style:family=\"{family}\">{properties_text}</style:style>"
            ),
        ));
        if family == "table-row" {
            self.rows.insert(name.to_string(), variant);
        } else {
            self.cols.insert(name.to_string(), variant);
        }
        Ok(())
    }

    /// The variants to add are those the document does not define already (a file that went through
    /// an earlier deletion has them).
    fn finish(&mut self, defined: &std::collections::HashSet<String>) {
        for (name, xml) in std::mem::take(&mut self.variants) {
            if !defined.contains(&name) {
                self.xml.push_str(&xml);
            }
        }
    }
}

/// The sizes of the rows and columns of each sheet of a `content.xml`, by the name of the sheet:
/// from the row and column styles in the automatic styles and the rows and columns that use them
/// (repeated ones, and those that are hidden or filtered, which count for nothing).
pub(super) fn scan_extents(
    input: impl std::io::BufRead,
) -> Result<(
    std::collections::HashMap<String, SheetExtents>,
    NoBreakStyles,
)> {
    use std::collections::HashMap;
    let mut no_break = NoBreakStyles::default();
    let mut defined: std::collections::HashSet<String> = std::collections::HashSet::new();
    // The default sizes: a row of 12.8 points and a column of 2.258 cm, those of `LibreOffice`.
    let (default_row, default_col) = (256 * 635, 1280 * 635);
    let mut row_height: HashMap<String, i64> = HashMap::new();
    let mut column_width: HashMap<String, i64> = HashMap::new();
    let mut out = HashMap::new();
    let mut reader = quick_xml::Reader::from_reader(input);
    let mut buf = Vec::new();
    let (mut style, mut family) = (String::new(), String::new());
    let (mut depth, mut sheet) = (0u32, String::new());
    let (mut rows, mut cols) = (Sizes::new(default_row), Sizes::new(default_col));
    let (mut row, mut col) = (0u32, 0u32);
    let repeated = |e: &BytesStart, key: &str| -> u32 {
        attribute(e, key)
            .ok()
            .flatten()
            .and_then(|n| n.parse::<u32>().ok())
            .unwrap_or(1)
            .max(1)
    };
    let hidden = |e: &BytesStart| -> bool {
        attribute(e, "table:visibility")
            .ok()
            .flatten()
            .is_some_and(|v| v != "visible")
    };
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) => match e.name().as_ref() {
                "style:style" => {
                    style = attribute(&e, "style:name")?.unwrap_or_default();
                    defined.insert(style.clone());
                    family = attribute(&e, "style:family")?.unwrap_or_default();
                }
                "style:table-row-properties" if family == "table-row" => {
                    no_break.note(&style, "table-row", &e)?;
                    if let Some(h) = attribute(&e, "style:row-height")?.and_then(|h| length_emu(&h))
                    {
                        row_height.insert(style.clone(), in_twips(h));
                    }
                }
                "style:table-column-properties" if family == "table-column" => {
                    no_break.note(&style, "table-column", &e)?;
                    if let Some(w) =
                        attribute(&e, "style:column-width")?.and_then(|w| length_emu(&w))
                    {
                        column_width.insert(style.clone(), in_twips(w));
                    }
                }
                "table:table" => {
                    depth += 1;
                    if depth == 1 {
                        sheet = attribute(&e, "table:name")?.unwrap_or_default();
                        rows = Sizes::new(default_row);
                        cols = Sizes::new(default_col);
                        (row, col) = (0, 0);
                    }
                }
                "table:table-column" if depth == 1 => {
                    let n = repeated(&e, "table:number-columns-repeated");
                    let width = if hidden(&e) {
                        0
                    } else {
                        attribute(&e, "table:style-name")?
                            .and_then(|name| column_width.get(&name).copied())
                            .unwrap_or(default_col)
                    };
                    cols.set_run(col, col.saturating_add(n) - 1, width);
                    col = col.saturating_add(n);
                }
                "table:table-row" if depth == 1 => {
                    let n = repeated(&e, "table:number-rows-repeated");
                    let height = if hidden(&e) {
                        0
                    } else {
                        attribute(&e, "table:style-name")?
                            .and_then(|name| row_height.get(&name).copied())
                            .unwrap_or(default_row)
                    };
                    rows.set_run(row, row.saturating_add(n) - 1, height);
                    row = row.saturating_add(n);
                }
                _ => {}
            },
            Event::End(e) if e.name().as_ref() == "table:table" => {
                if depth == 1 {
                    out.insert(
                        std::mem::take(&mut sheet),
                        SheetExtents {
                            rows: std::mem::replace(&mut rows, Sizes::new(default_row)),
                            cols: std::mem::replace(&mut cols, Sizes::new(default_col)),
                        },
                    );
                }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    no_break.finish(&defined);
    Ok((out, no_break))
}

/// A shape placed for the rows and columns that are deleted: the cell that holds it now, and its
/// events with its anchor changed.
pub(super) struct Placed {
    pub row: u32,
    pub col: u32,
    pub events: Vec<Event<'static>>,
}

/// The shapes of the cell at `(row, col)`, one per top-level `draw:*` element of `events`, placed
/// as `LibreOffice` places a shape whose rows or columns are deleted (see [`super::shapes`]): the
/// cell it is in now (its own index when its cell goes), its end cell and the offsets in them. A
/// shape that lies wholly inside what is deleted is left out.
pub(super) fn place_shapes(
    events: &[Event<'static>],
    s: &SheetStructure,
    row: u32,
    col: u32,
) -> Vec<Placed> {
    let mut placed = Vec::new();
    let mut depth = 0u32;
    let mut current: Vec<Event<'static>> = Vec::new();
    for event in events {
        match event {
            Event::Start(_) => depth += 1,
            Event::End(_) => depth = depth.saturating_sub(1),
            _ => {}
        }
        let opens = matches!(event, Event::Start(e) | Event::Empty(e) if depth <= 1 && e.name().as_ref().starts_with("draw:"))
            && current.is_empty();
        if opens || !current.is_empty() {
            current.push(event.clone());
            let closes = matches!(event, Event::Empty(_) | Event::End(_)) && depth == 0;
            if closes {
                let shape = std::mem::take(&mut current);
                if let Some(p) = place_shape(&shape, s, row, col) {
                    placed.push(p);
                }
            }
        }
    }
    placed
}

fn place_shape(shape: &[Event<'static>], s: &SheetStructure, row: u32, col: u32) -> Option<Placed> {
    let (Event::Start(first) | Event::Empty(first)) = &shape[0] else {
        return None;
    };
    let name = first.name().as_ref().to_string();
    let number = |key: &str| {
        attribute(first, key)
            .ok()
            .flatten()
            .and_then(|v| length_emu(&v))
    };
    let end_cell = attribute(first, "table:end-cell-address")
        .ok()
        .flatten()
        .and_then(|a| parse_ref(&a).map(|r| (a, r)));
    // A sheet with no sizes known has those `LibreOffice` gives a new one: 12.8 points and 2.258 cm.
    let defaults;
    let extents = if let Some(extents) = s.extents.as_deref() {
        extents
    } else {
        defaults = SheetExtents {
            rows: Sizes::new(256 * 635),
            cols: Sizes::new(1280 * 635),
        };
        &defaults
    };
    let (rows, cols) = (&extents.rows, &extents.cols);
    let (x, y) = (number("svg:x").unwrap_or(0), number("svg:y").unwrap_or(0));
    let (width, height) = (
        number("svg:width").unwrap_or(0),
        number("svg:height").unwrap_or(0),
    );
    let end = end_cell
        .as_ref()
        .and_then(|(_, r)| match &r.reference.area {
            RefArea::Cell(c) => Some((c.row, c.col)),
            _ => None,
        });
    let (end_x, end_y) = (
        number("table:end-x").unwrap_or(0),
        number("table:end-y").unwrap_or(0),
    );
    let (mut row_span, mut col_span) = match end {
        Some((er, ec)) => (
            Span::of_ends(rows, row, y, er, end_y),
            Span::of_ends(cols, col, x, ec, end_x),
        ),
        None => (
            Span::of_size(rows, row, y, height),
            Span::of_size(cols, col, x, width),
        ),
    };
    let kind = Kind {
        resizes: end.is_some(),
        rectangle: matches!(
            name.as_str(),
            "draw:rect"
                | "draw:frame"
                | "draw:ellipse"
                | "draw:caption"
                | "draw:control"
                | "draw:image"
        ),
    };
    if !s.rows.is_empty() {
        let new = rows.without(&s.rows);
        row_span = shapes::deleted(
            row_span,
            rows,
            &new,
            &s.rows,
            &kind,
            col_span.start == col_span.end,
        )?;
    }
    if !s.cols.is_empty() {
        let new = cols.without(&s.cols);
        col_span = shapes::deleted(
            col_span,
            cols,
            &new,
            &s.cols,
            &kind,
            row_span.start == row_span.end,
        )?;
    }
    let resized = |element: &BytesStart<'static>| -> BytesStart<'static> {
        let mut out = element.clone();
        out.clear_attributes();
        for attribute in element.attributes().flatten() {
            let key = attribute.key.as_ref().to_string();
            // What did not change keeps the text it had.
            let changed = |new: i64, old: i64| (new != old).then(|| emu_length(new));
            let value = match key.as_str() {
                "svg:x" => changed(col_span.start_offset, x),
                "svg:y" => changed(row_span.start_offset, y),
                "svg:width" if end.is_some() => changed(col_span.size, width),
                "svg:height" if end.is_some() => changed(row_span.size, height),
                "table:end-x" => changed(col_span.end_offset, end_x),
                "table:end-y" => changed(row_span.end_offset, end_y),
                "table:end-cell-address" => end_cell.as_ref().and_then(|(_, r)| {
                    let RefArea::Cell(cell) = &r.reference.area else {
                        return None;
                    };
                    let mut moved = *cell;
                    (moved.row, moved.col) = (row_span.end, col_span.end);
                    Some(print_ref(&OdsRef {
                        sheet_abs: r.sheet_abs,
                        repeats_sheet: r.repeats_sheet,
                        last_sheet: r.last_sheet.clone(),
                        reference: Reference {
                            sheet: r.reference.sheet.clone(),
                            area: RefArea::Cell(moved),
                        },
                    }))
                }),
                _ => None,
            };
            match value {
                Some(text) => out.push_attribute((key.as_str(), text.as_str())),
                None => out.push_attribute(attribute),
            }
        }
        out
    };
    let events = shape
        .iter()
        .enumerate()
        .map(|(i, event)| match (i, event) {
            (0, Event::Start(e)) => Event::Start(resized(e)),
            (0, Event::Empty(e)) => Event::Empty(resized(e)),
            _ => event.clone(),
        })
        .collect();
    Some(Placed {
        row: row_span.start,
        col: col_span.start,
        events,
    })
}

#[cfg(test)]
mod tests {
    use super::super::structure::StructureEdits;
    use super::*;
    use crate::write::ValueScope;
    use std::collections::HashSet;

    fn view(sheet: &str, spec: &str) -> SheetStructure<'static> {
        let scopes: Vec<ValueScope> = vec![spec.parse().unwrap()];
        let plan = Box::leak(Box::new(
            StructureEdits::new(
                Some(&[sheet.to_string(), "T".to_string()]),
                &HashSet::new(),
                &scopes,
                &[],
            )
            .unwrap(),
        ));
        let edits = Box::leak(Box::new(plan.edits()));
        SheetStructure {
            plan,
            edits,
            sheet: sheet.to_string(),
            rows: plan.rows_of(sheet),
            cols: plan.cols_of(sheet),
            recalculate: false,
            extents: None,
        }
    }

    fn rows(sheet: &str, spec: &str, formula: &str) -> Option<String> {
        let scopes: Vec<ValueScope> = vec![spec.parse().unwrap()];
        let plan = StructureEdits::new(
            Some(&[sheet.to_string(), "F".to_string()]),
            &HashSet::new(),
            &scopes,
            &[],
        )
        .unwrap();
        let edits = plan.edits();
        let structure = SheetStructure {
            plan: &plan,
            edits: &edits,
            sheet: "F".to_string(),
            rows: plan.rows_of("F"),
            cols: plan.cols_of("F"),
            recalculate: false,
            extents: None,
        };
        rewrite_expression(&structure, formula).unwrap()
    }

    #[test]
    fn ods_references_follow_the_deletion_as_libreoffice_writes_them() {
        // `T!4:5` deleted, read from `F`: what LibreOffice wrote after `Rows.removeByIndex`.
        let t = |f: &str| rows("T", "T!4:5", f);
        assert_eq!(t("[$T.A2]"), None);
        assert_eq!(t("[$T.A4]").as_deref(), Some("[#REF!]"));
        assert_eq!(t("[$T.A6]").as_deref(), Some("[$T.A4]"));
        assert_eq!(t("SUM([$T.A2:.A4])").as_deref(), Some("SUM([$T.A2:.A3])"));
        assert_eq!(t("SUM([$T.A4:.A5])").as_deref(), Some("SUM([#REF!])"));
        assert_eq!(t("SUM([$T.A5:.A6])").as_deref(), Some("SUM([$T.A4:.A4])"));
        assert_eq!(t("SUM([$T.A:.A])"), None);
        assert_eq!(t("SUM([$T.3:.6])").as_deref(), Some("SUM([$T.3:.4])"));
        assert_eq!(t("[$T.$A$6]").as_deref(), Some("[$T.$A$4]"));
        assert_eq!(t("[$T.A6]+[.A6]").as_deref(), Some("[$T.A4]+[.A6]"));
        // Text, names and external references are not touched.
        assert_eq!(
            t("\"[$T.A6]\"&[$T.A6]").as_deref(),
            Some("\"[$T.A6]\"&[$T.A4]")
        );
        assert_eq!(t("['file:///x.ods'#$T.A6]"), None);
        // A whole row marked absolute, and a list of ranges whose sheet name has a space.
        assert_eq!(t("[$T.$3:.$6]").as_deref(), Some("[$T.$3:.$4]"));
        assert_eq!(
            adjust_range_list(&view("F", "F!2"), "'My Sheet'.A3:.A5 F.A3:F.B4")
                .unwrap()
                .as_deref(),
            Some("'My Sheet'.A3:.A5 F.A2:F.B3")
        );
        // A reference without sheet reads the sheet of the formula.
        assert_eq!(
            rows("F", "F!2", "[.A3]+[$T.A3]").as_deref(),
            Some("[.A2]+[$T.A3]")
        );
    }

    #[test]
    fn a_reference_across_sheets_is_refused() {
        let scopes: Vec<ValueScope> = vec!["T!4:5".parse().unwrap()];
        let plan = StructureEdits::new(None, &HashSet::new(), &scopes, &[]).unwrap();
        let edits = plan.edits();
        let structure = SheetStructure {
            plan: &plan,
            edits: &edits,
            sheet: "F".to_string(),
            rows: plan.rows_of("F"),
            cols: plan.cols_of("F"),
            recalculate: false,
            extents: None,
        };
        for formula in ["SUM([$T.A2:$V.A4])", "[$T.A1:$V.B2]+1"] {
            let error = rewrite_expression(&structure, formula).unwrap_err();
            let error = format!("{error}");
            assert!(error.contains("3D reference"), "{error}");
            assert!(error.contains("issues/75"), "{error}");
        }
        // A range inside one sheet is not one, in either way of writing the second end.
        assert!(rewrite_expression(&structure, "SUM([$T.A2:.A4])").is_ok());
        assert!(rewrite_expression(&structure, "SUM([$T.A2:$T.A4])").is_ok());
    }
}

//! References of an ODS document under a structural plan: the formulas of the cells
//! (`of:=SUM([$T.A2:.A3])`), the named expressions and the attributes that hold ranges.
//!
//! The semantics are those of `sheetrs_formula::edit`, which reads Excel's A1 syntax; this module
//! reads and writes the notation of ODF (`[$Sheet.A1:.B2]`) around it.

use super::attribute;
use super::xlsx_structure::SheetStructure;
use anyhow::{Result, bail};
use quick_xml::events::{BytesStart, Event};
use sheetrs_formula::addr::same_name;
use sheetrs_formula::addr::{col_index, col_name, needs_quotes};
use sheetrs_formula::ast::{CellAddr, ColAddr, Expr, RefArea, Reference, RowAddr};
use sheetrs_formula::edit::{ThreeD, apply, apply_three_d};

/// A reference as an ODS formula writes it, inside `[` `]`.
struct OdsRef {
    /// The sheet is marked absolute (`$Sheet.`).
    sheet_abs: bool,
    /// The second end of a range repeats the sheet (`S.A1:S.B2`), as an attribute writes it,
    /// instead of leaving it out (`$S.A1:.B2`), as a formula does.
    repeats_sheet: bool,
    /// The last sheet of a 3D reference (`$A.A1:$C.B2`), when it is not the first.
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
fn adjusted(s: &SheetStructure, content: &str) -> Option<Option<String>> {
    let parsed = parse_ref(content)?;
    if let (Some(first), Some(last)) = (&parsed.reference.sheet, &parsed.last_sheet) {
        return adjusted_three_d(s, &parsed, first, last);
    }
    let mut node = Expr::Ref(parsed.reference.clone());
    let mut changed = false;
    for edit in s.edits {
        changed |= apply(&mut node, edit, &s.sheet);
    }
    if !changed {
        return None;
    }
    Some(match node {
        Expr::Ref(reference) => Some(print_ref(&OdsRef {
            sheet_abs: parsed.sheet_abs,
            repeats_sheet: parsed.repeats_sheet,
            last_sheet: None,
            reference,
        })),
        _ => None,
    })
}

/// A 3D reference of `first` to `last`, by the rules of
/// [`sheetrs_formula::edit::apply_three_d`].
fn adjusted_three_d(
    s: &SheetStructure,
    parsed: &OdsRef,
    first: &str,
    last: &str,
) -> Option<Option<String>> {
    let removed = |name: &str| s.plan.removes_sheet(name);
    let book = s.plan.workbook(&removed);
    let (mut first, mut last, mut area) =
        (first.to_string(), last.to_string(), parsed.reference.area);
    let mut changed = false;
    for edit in s.edits {
        match apply_three_d(&first, &last, &area, edit, &book) {
            ThreeD::Same => {}
            ThreeD::Now {
                first: f,
                last: l,
                area: a,
            } => {
                (first, last, area) = (f, l, a);
                changed = true;
            }
            ThreeD::CellsGone | ThreeD::SheetsGone => return Some(None),
        }
    }
    changed.then(|| {
        Some(print_ref(&OdsRef {
            sheet_abs: parsed.sheet_abs,
            repeats_sheet: parsed.repeats_sheet,
            last_sheet: (!same_name(&first, &last)).then_some(last),
            reference: Reference {
                sheet: Some(first),
                area,
            },
        }))
    })
}

/// The text of the ODS formula (or expression) `text` once the plan is applied, `None` when it
/// does not change. References in brackets are adjusted; one that goes is `[#REF!]`.
pub(super) fn rewrite_expression(s: &SheetStructure, text: &str) -> Option<String> {
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
                if let Some(new) = adjusted(s, &text[i + 1..j]) {
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
    changed.then_some(out)
}

/// The formula of a `table:formula` attribute (`of:=…`) once the plan is applied.
pub(super) fn rewrite_cell_formula(s: &SheetStructure, formula: &str) -> Result<Option<String>> {
    if !s.changes_layout() && !s.plan.may_reach_edited_sheet(formula) {
        return Ok(None);
    }
    let Some(body) = formula.strip_prefix("of:=") else {
        bail!("the formula '{formula}' is not in the OpenFormula syntax (of:=)");
    };
    Ok(rewrite_expression(s, body).map(|new| format!("of:={new}")))
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
pub(super) fn adjust_range_list(s: &SheetStructure, value: &str) -> Option<String> {
    let kept: Vec<String> = split_ranges(value)
        .into_iter()
        .filter(|range| !range.is_empty())
        .filter_map(|range| match adjusted(s, range) {
            None => Some(range.to_string()),
            Some(new) => new,
        })
        .collect();
    (!kept.is_empty()).then(|| kept.join(" "))
}

/// The shapes of a cell with the `table:end-cell-address` of each moved as LibreOffice moves the
/// cell of a shape that resizes with its cells: to the place of that cell, or to its own index
/// when the cell goes.
pub(super) fn move_end_anchors<'a>(events: &[Event<'a>], s: &SheetStructure) -> Vec<Event<'a>> {
    let moved = |element: &BytesStart<'a>| -> BytesStart<'a> {
        let Ok(Some(address)) = attribute(element, "table:end-cell-address") else {
            return element.clone();
        };
        let Some(OdsRef {
            sheet_abs,
            repeats_sheet,
            last_sheet,
            reference:
                Reference {
                    sheet,
                    area: RefArea::Cell(mut end),
                },
        }) = parse_ref(&address)
        else {
            return element.clone();
        };
        (end.row, end.col) = (s.rows.anchored_at(end.row), s.cols.anchored_at(end.col));
        let text = print_ref(&OdsRef {
            sheet_abs,
            repeats_sheet,
            last_sheet,
            reference: Reference {
                sheet,
                area: RefArea::Cell(end),
            },
        });
        let mut out = element.clone();
        out.clear_attributes();
        for attribute in element.attributes().flatten() {
            if attribute.key.as_ref() == "table:end-cell-address" {
                out.push_attribute(("table:end-cell-address", text.as_str()));
            } else {
                out.push_attribute(attribute);
            }
        }
        out
    };
    events
        .iter()
        .map(|event| match event {
            Event::Start(e) if e.name().as_ref().starts_with("draw:") => Event::Start(moved(e)),
            Event::Empty(e) if e.name().as_ref().starts_with("draw:") => Event::Empty(moved(e)),
            other => other.clone(),
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::super::structure::StructureEdits;
    use super::*;
    use crate::writer::ValueScope;
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
        };
        rewrite_expression(&structure, formula)
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
            adjust_range_list(&view("F", "F!2"), "'My Sheet'.A3:.A5 F.A3:F.B4").as_deref(),
            Some("'My Sheet'.A3:.A5 F.A2:F.B3")
        );
        // A reference without sheet reads the sheet of the formula.
        assert_eq!(
            rows("F", "F!2", "[.A3]+[$T.A3]").as_deref(),
            Some("[.A2]+[$T.A3]")
        );
    }
}

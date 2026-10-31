//! Structural edits on references: what a formula becomes when rows or columns
//! are inserted or deleted, or a sheet is renamed or deleted, as the
//! spreadsheet applications do it.
//!
//! A reference to deleted cells becomes `Sheet!#REF!` (`#REF!` for a deleted sheet); a range
//! shrinks or grows around
//! the edit; absolute markers are irrelevant (both kinds move).

use crate::addr::needs_quotes;
use crate::addr::{MAX_COLS, MAX_ROWS, same_name};
use crate::ast::{BinOp, Expr, RefArea, Reference};
use crate::print::sheet_text;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Edit<'a> {
    /// Insert `count` rows before the zero-based row `at` of `sheet`.
    InsertRows {
        sheet: &'a str,
        at: u32,
        count: u32,
    },
    DeleteRows {
        sheet: &'a str,
        at: u32,
        count: u32,
    },
    InsertCols {
        sheet: &'a str,
        at: u32,
        count: u32,
    },
    DeleteCols {
        sheet: &'a str,
        at: u32,
        count: u32,
    },
    DeleteSheet {
        sheet: &'a str,
    },
    RenameSheet {
        from: &'a str,
        to: &'a str,
    },
}

/// Which axis an edit moves and by how much.
#[derive(Clone, Copy)]
enum Shift {
    Insert { at: u32, count: u32 },
    Delete { at: u32, count: u32 },
}

/// New position of a single index, `None` when deleted or out of range.
fn point(i: u32, s: Shift, limit: u32) -> Option<u32> {
    match s {
        Shift::Insert { at, count } => {
            let n = if i >= at { i + count } else { i };
            (n < limit).then_some(n)
        }
        Shift::Delete { at, count } => {
            if i < at {
                Some(i)
            } else if i >= at + count {
                Some(i - count)
            } else {
                None
            }
        }
    }
}

/// New bounds of `lo..=hi`, `None` when nothing of it remains.
fn span(lo: u32, hi: u32, s: Shift, limit: u32) -> Option<(u32, u32)> {
    match s {
        Shift::Insert { at, count } => {
            let l = if lo >= at { lo + count } else { lo };
            let h = if hi >= at { hi + count } else { hi };
            (l < limit).then_some((l, h.min(limit - 1)))
        }
        Shift::Delete { at, count } => {
            let end = at + count;
            let l = if lo < at {
                lo
            } else if lo >= end {
                lo - count
            } else {
                at
            };
            let h = if hi < at {
                i64::from(hi)
            } else if hi >= end {
                i64::from(hi - count)
            } else {
                i64::from(at) - 1
            };
            (h >= i64::from(l)).then_some((l, h as u32))
        }
    }
}

fn ordered(a: u32, b: u32) -> (u32, u32, bool) {
    if a <= b { (a, b, false) } else { (b, a, true) }
}

/// Apply `shift` along rows (`rows == true`) or columns to an area.
fn shift_area(area: &RefArea, s: Shift, rows: bool) -> Option<RefArea> {
    let (limit, _) = if rows { (MAX_ROWS, ()) } else { (MAX_COLS, ()) };
    Some(match *area {
        RefArea::Cell(mut c) => {
            if rows {
                c.row = point(c.row, s, limit)?;
            } else {
                c.col = point(c.col, s, limit)?;
            }
            RefArea::Cell(c)
        }
        RefArea::Range(mut a, mut b) => {
            let (x, y) = if rows { (a.row, b.row) } else { (a.col, b.col) };
            let (lo, hi, swapped) = ordered(x, y);
            let (l, h) = span(lo, hi, s, limit)?;
            let (nx, ny) = if swapped { (h, l) } else { (l, h) };
            if rows {
                (a.row, b.row) = (nx, ny);
            } else {
                (a.col, b.col) = (nx, ny);
            }
            RefArea::Range(a, b)
        }
        RefArea::Cols(mut a, mut b) if !rows => {
            let (lo, hi, swapped) = ordered(a.col, b.col);
            let (l, h) = span(lo, hi, s, limit)?;
            (a.col, b.col) = if swapped { (h, l) } else { (l, h) };
            RefArea::Cols(a, b)
        }
        RefArea::Rows(mut a, mut b) if rows => {
            let (lo, hi, swapped) = ordered(a.row, b.row);
            let (l, h) = span(lo, hi, s, limit)?;
            (a.row, b.row) = if swapped { (h, l) } else { (l, h) };
            RefArea::Rows(a, b)
        }
        // A whole-column reference does not move when rows change, and the
        // other way round.
        other => other,
    })
}

/// Rewrite the references of `expr`. `home_sheet` is the sheet that holds the
/// formula (unqualified references point there). Returns whether anything
/// changed.
pub fn apply(expr: &mut Expr, edit: &Edit<'_>, home_sheet: &str) -> bool {
    let mut changed = false;
    expr.walk_mut(&mut |e| {
        let Expr::Ref(r) = &*e else { return };
        let mut r = r.clone();
        let target = r.sheet.clone().unwrap_or_else(|| home_sheet.to_string());
        let target = target.as_str();
        match *edit {
            Edit::RenameSheet { from, to } => {
                if r.sheet.as_deref().is_some_and(|s| same_name(s, from)) {
                    r.sheet = Some(to.to_string());
                    *e = Expr::Ref(r);
                    changed = true;
                }
            }
            Edit::DeleteSheet { sheet } => {
                if r.sheet.as_deref().is_some_and(|s| same_name(s, sheet)) {
                    *e = Expr::Error("#REF!".into());
                    changed = true;
                }
            }
            Edit::InsertRows { sheet, at, count } | Edit::DeleteRows { sheet, at, count }
                if same_name(target, sheet) =>
            {
                let s = if matches!(edit, Edit::InsertRows { .. }) {
                    Shift::Insert { at, count }
                } else {
                    Shift::Delete { at, count }
                };
                changed |= replace(e, r, s, true, target);
            }
            Edit::InsertCols { sheet, at, count } | Edit::DeleteCols { sheet, at, count }
                if same_name(target, sheet) =>
            {
                let s = if matches!(edit, Edit::InsertCols { .. }) {
                    Shift::Insert { at, count }
                } else {
                    Shift::Delete { at, count }
                };
                changed |= replace(e, r, s, false, target);
            }
            _ => {}
        }
    });
    changed
}

/// Rewrite the reference `r` to a sheet named `target` after `s`. Cells that are gone leave
/// `Sheet!#REF!`, the form that LibreOffice reads as `#REF!` (a bare one is an unknown name there)
/// and Excel writes.
fn replace(e: &mut Expr, r: Reference, s: Shift, rows: bool, target: &str) -> bool {
    match shift_area(&r.area, s, rows) {
        Some(area) if area == r.area => false,
        Some(area) => {
            *e = Expr::Ref(Reference {
                sheet: r.sheet,
                area,
            });
            true
        }
        None => {
            *e = Expr::Error(format!("{}!#REF!", sheet_text(target)));
            true
        }
    }
}

/// The sheets of a workbook, in order: the span of a 3D reference (`Sheet1:Sheet3!A1`) is
/// the sheets between its two ends.
#[derive(Clone, Copy)]
pub struct Workbook<'a> {
    pub order: &'a [String],
    /// Whether the sheet is one the command removes.
    pub removed: &'a dyn Fn(&str) -> bool,
}

impl Workbook<'_> {
    fn position(&self, name: &str) -> Option<usize> {
        self.order.iter().position(|s| same_name(s, name))
    }
}

/// What a 3D reference becomes after an edit.
#[derive(Debug, Clone, PartialEq)]
pub enum ThreeD {
    Same,
    /// The cells it names are gone.
    CellsGone,
    /// Every sheet of its span is gone.
    SheetsGone,
    Now {
        first: String,
        last: String,
        area: RefArea,
    },
}

/// The bounds of `area` along rows (or columns), `None` when it does not extend along that axis.
fn bounds(area: &RefArea, rows: bool) -> Option<(u32, u32)> {
    let (x, y) = match *area {
        RefArea::Cell(c) => (
            if rows { c.row } else { c.col },
            if rows { c.row } else { c.col },
        ),
        RefArea::Range(a, b) if rows => (a.row, b.row),
        RefArea::Range(a, b) => (a.col, b.col),
        RefArea::Rows(a, b) if rows => (a.row, b.row),
        RefArea::Cols(a, b) if !rows => (a.col, b.col),
        _ => return None,
    };
    Some((x.min(y), x.max(y)))
}

/// The 3D reference `first:last!area` after `edit`, as LibreOffice does it:
///
/// - Rows or columns deleted on a sheet of the span change the reference only if they touch the
///   rows or columns it names; deleted ones before it do not move it. Otherwise the reference
///   follows the edit as a reference to one sheet does. A sheet outside the span changes nothing.
/// - A removed sheet at an end of the span leaves the span to the next kept sheet; one in the
///   middle changes nothing; the span of removed sheets only is gone.
///
/// The renaming of a sheet renames the end of the span that has its name. Insertions are not
/// applied to a 3D reference: no tool of this workspace inserts rows or columns.
pub fn apply_three_d(
    first: &str,
    last: &str,
    area: &RefArea,
    edit: &Edit<'_>,
    book: &Workbook<'_>,
) -> ThreeD {
    let (Some(a), Some(b)) = (book.position(first), book.position(last)) else {
        return ThreeD::Same;
    };
    let (a, b) = (a.min(b), a.max(b));
    let (sheet, shift, rows) = match *edit {
        Edit::DeleteRows { sheet, at, count } => (sheet, Shift::Delete { at, count }, true),
        Edit::DeleteCols { sheet, at, count } => (sheet, Shift::Delete { at, count }, false),
        Edit::DeleteSheet { sheet } => {
            if !book.position(sheet).is_some_and(|x| x == a || x == b) {
                return ThreeD::Same;
            }
            let kept: Vec<&String> = book.order[a..=b]
                .iter()
                .filter(|name| !(book.removed)(name))
                .collect();
            return match (kept.first(), kept.last()) {
                (Some(f), Some(l)) => ThreeD::Now {
                    first: (*f).clone(),
                    last: (*l).clone(),
                    area: *area,
                },
                _ => ThreeD::SheetsGone,
            };
        }
        Edit::RenameSheet { from, to } => {
            let renamed = |name: &str| if same_name(name, from) { to } else { name }.to_string();
            return if same_name(first, from) || same_name(last, from) {
                ThreeD::Now {
                    first: renamed(first),
                    last: renamed(last),
                    area: *area,
                }
            } else {
                ThreeD::Same
            };
        }
        _ => return ThreeD::Same,
    };
    if !book.position(sheet).is_some_and(|x| (a..=b).contains(&x)) {
        return ThreeD::Same;
    }
    let Shift::Delete { at, count } = shift else {
        return ThreeD::Same;
    };
    let Some((lo, hi)) = bounds(area, rows) else {
        return ThreeD::Same;
    };
    if hi < at || lo >= at + count {
        return ThreeD::Same;
    }
    match shift_area(area, shift, rows) {
        Some(new) if new == *area => ThreeD::Same,
        Some(area) => ThreeD::Now {
            first: first.to_string(),
            last: last.to_string(),
            area,
        },
        None => ThreeD::CellsGone,
    }
}

/// The qualifier of a 3D reference as it is written before `!`.
fn three_d_sheets(first: &str, last: &str) -> String {
    if same_name(first, last) {
        sheet_text(first)
    } else if needs_quotes(first) || needs_quotes(last) {
        format!(
            "'{}:{}'",
            first.replace('\'', "''"),
            last.replace('\'', "''")
        )
    } else {
        format!("{first}:{last}")
    }
}

impl crate::parse::Formula {
    /// The formula text after `edit`, or `None` when it does not change.
    ///
    /// Only the references that change are replaced, so whitespace, letter case,
    /// function-name spelling and everything else in the source stay exactly as
    /// written.
    pub fn rewrite(&self, edit: &Edit<'_>, home_sheet: &str) -> Option<String> {
        self.rewrite_all(std::slice::from_ref(edit), home_sheet, None)
    }

    /// The formula text after every edit of `edits`, applied to each reference in order, or
    /// `None` when it does not change. The edits are written in the coordinates the formula has
    /// now: to apply several deletions of one sheet at once, list them from the last row (or
    /// column) to the first, so each leaves the positions of the next one as they were.
    ///
    /// With the sheets of the `book`, a 3D reference follows the rules of [`apply_three_d`].
    /// Without them 3D references are not recognised: the last reference of one is read as a
    /// reference to its own sheet.
    pub fn rewrite_all(
        &self,
        edits: &[Edit<'_>],
        home_sheet: &str,
        book: Option<&Workbook<'_>>,
    ) -> Option<String> {
        let refs = self.expr.references();
        let three_d = book.map(|book| self.three_d_ends(book)).unwrap_or_default();
        let mut out = String::with_capacity(self.source.len());
        let mut last = 0;
        let mut changed = false;
        for (i, (r, span)) in refs.into_iter().zip(&self.ref_spans).enumerate() {
            let three = three_d.iter().find(|(at, ..)| *at == i);
            let (start, replacement) = match (three, book) {
                (Some((_, first, last_sheet, from)), Some(book)) => {
                    let (mut f, mut l, mut area) = (first.clone(), last_sheet.clone(), r.area);
                    let mut now = ThreeD::Same;
                    for edit in edits {
                        match apply_three_d(&f, &l, &area, edit, book) {
                            ThreeD::Same => {}
                            ThreeD::Now {
                                first,
                                last,
                                area: moved,
                            } => {
                                (f, l, area) = (first, last, moved);
                                now = ThreeD::Now {
                                    first: f.clone(),
                                    last: l.clone(),
                                    area,
                                };
                            }
                            gone => {
                                now = gone;
                                break;
                            }
                        }
                    }
                    let text = match now {
                        ThreeD::Same => continue,
                        ThreeD::CellsGone => format!("{}!#REF!", sheet_text(&f)),
                        ThreeD::SheetsGone => "#REF!".to_string(),
                        ThreeD::Now { first, last, area } => format!(
                            "{}!{}",
                            three_d_sheets(&first, &last),
                            Reference { sheet: None, area }
                        ),
                    };
                    (*from, text)
                }
                _ => {
                    let mut node = Expr::Ref(r.clone());
                    let mut touched = false;
                    for edit in edits {
                        touched |= apply(&mut node, edit, home_sheet);
                    }
                    if !touched {
                        continue;
                    }
                    (span.start, node.to_string())
                }
            };
            out.push_str(&self.source[last..start]);
            out.push_str(&replacement);
            last = span.end;
            changed = true;
        }
        out.push_str(&self.source[last..]);
        changed.then_some(out)
    }

    /// The 3D references of the formula: the index in [`Expr::references`] of the reference that
    /// ends each (`Sheet1:Sheet3!A1` reads as the name `Sheet1` and the reference `Sheet3!A1`, a
    /// quoted `'Sheet 1:Sheet 3'!A1` as one sheet), their first and last sheet, and where their
    /// text starts.
    fn three_d_ends(&self, book: &Workbook<'_>) -> Vec<(usize, String, String, usize)> {
        let refs = self.expr.references();
        let mut ends = Vec::new();
        let mut named: Vec<(*const Reference, String)> = Vec::new();
        self.expr.walk(&mut |e| {
            if let Expr::Binary(BinOp::Range, a, b) = e
                && let (Expr::Name(first), Expr::Ref(r)) = (&**a, &**b)
                && r.sheet.is_some()
                && book.position(first).is_some()
            {
                named.push((r, first.clone()));
            }
        });
        for (i, (r, span)) in refs.into_iter().zip(&self.ref_spans).enumerate() {
            if let Some((_, first)) = named.iter().find(|(p, _)| std::ptr::eq(*p, r)) {
                let prefix = format!("{first}:");
                let from = span.start.checked_sub(prefix.len());
                if let (Some(from), Some(last)) = (from, &r.sheet)
                    && self.source.get(from..span.start) == Some(prefix.as_str())
                {
                    ends.push((i, first.clone(), last.clone(), from));
                }
            } else if let Some((first, last)) = r.sheet.as_deref().and_then(|s| s.split_once(':')) {
                ends.push((i, first.to_string(), last.to_string(), span.start));
            }
        }
        ends
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parse::parse;

    fn run(f: &str, edit: Edit<'_>, home: &str) -> String {
        let mut e = parse(f).unwrap();
        apply(&mut e, &edit, home);
        e.to_string()
    }

    fn del_rows(f: &str, at: u32, n: u32) -> String {
        run(
            f,
            Edit::DeleteRows {
                sheet: "S",
                at,
                count: n,
            },
            "S",
        )
    }

    #[test]
    fn deleting_rows() {
        // Row 3 (index 2) deleted.
        assert_eq!(del_rows("A1+A3+A4", 2, 1), "A1+S!#REF!+A3");
        assert_eq!(del_rows("SUM(A1:A5)", 2, 1), "SUM(A1:A4)");
        assert_eq!(del_rows("SUM(A3:A5)", 2, 1), "SUM(A3:A4)");
        assert_eq!(del_rows("SUM(A1:A3)", 2, 1), "SUM(A1:A2)");
        assert_eq!(del_rows("SUM(A3:A3)", 2, 1), "SUM(S!#REF!)");
        assert_eq!(del_rows("SUM(A2:A4)", 1, 3), "SUM(S!#REF!)");
        assert_eq!(del_rows("A$4+$B5", 0, 2), "A$2+$B3");
        assert_eq!(del_rows("A:A+3:4", 0, 1), "A:A+2:3");
        assert_eq!(del_rows("SUM(A5:A1)", 2, 1), "SUM(A4:A1)");
    }

    #[test]
    fn inserting_rows_and_columns() {
        let ins = |f, at, n| {
            run(
                f,
                Edit::InsertRows {
                    sheet: "S",
                    at,
                    count: n,
                },
                "S",
            )
        };
        assert_eq!(ins("A1+A3", 2, 2), "A1+A5");
        assert_eq!(
            ins("SUM(A1:A3)", 2, 1),
            "SUM(A1:A4)",
            "an insertion inside grows the range"
        );
        assert_eq!(ins("SUM(A1:A3)", 3, 1), "SUM(A1:A3)");
        let cols = |f, at, n| {
            run(
                f,
                Edit::DeleteCols {
                    sheet: "S",
                    at,
                    count: n,
                },
                "S",
            )
        };
        assert_eq!(cols("B1+C1+D1", 2, 1), "B1+S!#REF!+C1");
        assert_eq!(cols("SUM(A1:D1)", 1, 2), "SUM(A1:B1)");
        assert_eq!(cols("B:D", 0, 1), "A:C");
        assert_eq!(
            run(
                "XFD1",
                Edit::InsertCols {
                    sheet: "S",
                    at: 0,
                    count: 1
                },
                "S"
            ),
            "S!#REF!"
        );
    }

    #[test]
    fn only_the_edited_sheet_moves() {
        let e = Edit::DeleteRows {
            sheet: "Data",
            at: 0,
            count: 1,
        };
        // Unqualified reference on another sheet: untouched.
        assert_eq!(
            run("A5+Data!A5+'Data'!A1", e, "Other"),
            "A5+Data!A4+Data!#REF!"
        );
        // Unqualified reference on the edited sheet itself moves.
        assert_eq!(run("A5", e, "data"), "A4");
    }

    #[test]
    fn rewriting_keeps_the_source_text() {
        use crate::parse::parse_spanned;
        let f = parse_spanned("=  sum( A1 : A5 ,  b3 ) + IF( x ,  Data!A9 ,'Q 1'!C3 )").unwrap();
        // Delete row 3 of `S`, the home sheet: A5 → A4, b3 → #REF!; Data and
        // 'Q 1' are other sheets and stay.
        let e = Edit::DeleteRows {
            sheet: "S",
            at: 2,
            count: 1,
        };
        assert_eq!(
            f.rewrite(&e, "S").unwrap(),
            "=  sum( A1 : A4 ,  S!#REF! ) + IF( x ,  Data!A9 ,'Q 1'!C3 )"
        );
        // Nothing to change: no new text.
        assert_eq!(
            f.rewrite(
                &Edit::DeleteRows {
                    sheet: "Elsewhere",
                    at: 0,
                    count: 1
                },
                "S"
            ),
            None
        );
        // A rename touches the qualifier only.
        let r = Edit::RenameSheet {
            from: "data",
            to: "Data 2",
        };
        assert_eq!(
            f.rewrite(&r, "S").unwrap(),
            "=  sum( A1 : A5 ,  b3 ) + IF( x ,  'Data 2'!A9 ,'Q 1'!C3 )"
        );
        // Spans follow the source, with or without the leading `=`.
        let g = parse_spanned("A1+B2").unwrap();
        assert_eq!(&g.source[g.ref_spans[1].clone()], "B2");
        let h = parse_spanned("=A1+B2").unwrap();
        assert_eq!(&h.source[h.ref_spans[1].clone()], "B2");
    }

    /// Several deletions of one sheet, in the coordinates of the file, as LibreOffice gives them.
    #[test]
    fn several_edits_apply_to_the_coordinates_of_the_input() {
        use crate::parse::parse_spanned;
        let rows = |f: &str, bands: &[(u32, u32)]| {
            // Last band first: each leaves the positions of the next as they were.
            let edits: Vec<_> = bands
                .iter()
                .rev()
                .map(|&(at, count)| Edit::DeleteRows {
                    sheet: "S",
                    at,
                    count,
                })
                .collect();
            parse_spanned(f).unwrap().rewrite_all(&edits, "S", None)
        };
        // Rows 2-3 and 6 deleted (1-based 3-4 and 7): A5 moves up two, A8 up three.
        assert_eq!(
            rows("=A5+A8+A3", &[(2, 2), (6, 1)]).as_deref(),
            Some("=A3+A5+S!#REF!")
        );
        assert_eq!(
            rows("=SUM(A1:A10)", &[(2, 2), (6, 1)]).as_deref(),
            Some("=SUM(A1:A7)")
        );
        assert_eq!(rows("=A1+A2", &[(2, 2)]), None);
    }

    /// The sheets `Alpha` to `Delta`, in that order (names that no cell has); `removed` are the ones the command removes.
    fn three_d(formula: &str, edits: &[Edit<'_>], removed: &[&str]) -> Option<String> {
        use crate::parse::parse_spanned;
        let order: Vec<String> = ["Alpha", "Beta", "Gamma", "Delta"]
            .map(String::from)
            .to_vec();
        let removes = |name: &str| removed.contains(&name);
        let book = Workbook {
            order: &order,
            removed: &removes,
        };
        parse_spanned(formula)
            .unwrap()
            .rewrite_all(edits, "F", Some(&book))
    }

    fn rows_of(sheet: &'static str, at: u32, count: u32) -> Edit<'static> {
        Edit::DeleteRows { sheet, at, count }
    }

    /// What LibreOffice 26.8 gave (`sheetrs-core/scripts/compare_structure_with_libreoffice.py`).
    #[test]
    fn a_3d_reference_changes_only_where_the_deletion_touches_it() {
        // A deletion before what it names, in a sheet of the span or out of it, changes nothing.
        assert_eq!(
            three_d("=SUM(Alpha:Gamma!A3:A5)", &[rows_of("Beta", 0, 2)], &[]),
            None
        );
        assert_eq!(
            three_d("=SUM(Alpha:Gamma!A3:A5)", &[rows_of("Delta", 2, 1)], &[]),
            None
        );
        // One that touches it shrinks it, wherever the sheet is in the span.
        for sheet in ["Alpha", "Beta", "Gamma"] {
            assert_eq!(
                three_d("=SUM(Alpha:Gamma!A3:A5)", &[rows_of(sheet, 3, 1)], &[]).as_deref(),
                Some("=SUM(Alpha:Gamma!A3:A4)"),
                "{sheet}"
            );
        }
        assert_eq!(
            three_d("=SUM(Alpha:Gamma!A3)", &[rows_of("Beta", 2, 1)], &[]).as_deref(),
            Some("=SUM(Alpha!#REF!)")
        );
        // A whole column does not move when rows go.
        assert_eq!(
            three_d("=SUM(Alpha:Gamma!A:A)", &[rows_of("Beta", 2, 1)], &[]),
            None
        );
    }

    #[test]
    fn a_removed_sheet_of_a_3d_span_passes_the_end_on() {
        let gone = |sheet| [Edit::DeleteSheet { sheet }];
        assert_eq!(
            three_d("=SUM(Alpha:Gamma!A1)", &gone("Alpha"), &["Alpha"]).as_deref(),
            Some("=SUM(Beta:Gamma!A1)")
        );
        assert_eq!(
            three_d("=SUM(Alpha:Gamma!A1)", &gone("Gamma"), &["Gamma"]).as_deref(),
            Some("=SUM(Alpha:Beta!A1)")
        );
        assert_eq!(
            three_d("=SUM(Alpha:Gamma!A1)", &gone("Beta"), &["Beta"]),
            None
        );
        assert_eq!(
            three_d("=SUM(Alpha:Beta!A1)", &gone("Alpha"), &["Alpha"]).as_deref(),
            Some("=SUM(Beta!A1)")
        );
        // Every sheet of the span goes, one after the other.
        let both = [
            Edit::DeleteSheet { sheet: "Alpha" },
            Edit::DeleteSheet { sheet: "Beta" },
        ];
        assert_eq!(
            three_d("=SUM(Alpha:Beta!A1)", &both, &["Alpha", "Beta"]).as_deref(),
            Some("=SUM(#REF!)")
        );
        // The end that moved is the one the next edit reads, and the one a lost cell names.
        let edits = [Edit::DeleteSheet { sheet: "Alpha" }, rows_of("Beta", 3, 1)];
        assert_eq!(
            three_d("=SUM(Alpha:Gamma!A2:A5)", &edits, &["Alpha"]).as_deref(),
            Some("=SUM(Beta:Gamma!A2:A4)")
        );
        let edits = [Edit::DeleteSheet { sheet: "Alpha" }, rows_of("Beta", 2, 1)];
        assert_eq!(
            three_d("=SUM(Alpha:Gamma!A3)", &edits, &["Alpha"]).as_deref(),
            Some("=SUM(Beta!#REF!)")
        );
    }

    #[test]
    fn a_renamed_sheet_renames_the_end_of_a_3d_span() {
        let rename = [Edit::RenameSheet {
            from: "gamma",
            to: "Last one",
        }];
        assert_eq!(
            three_d("=SUM(Alpha:Gamma!A1)", &rename, &[]).as_deref(),
            Some("=SUM('Alpha:Last one'!A1)")
        );
    }

    #[test]
    fn sheets_renamed_and_deleted() {
        let r = Edit::RenameSheet {
            from: "Old",
            to: "New name",
        };
        assert_eq!(
            run("Old!A1+SUM(old!B1:B2)", r, "X"),
            "'New name'!A1+SUM('New name'!B1:B2)"
        );
        let d = Edit::DeleteSheet { sheet: "Gone" };
        assert_eq!(run("Gone!A1+B1", d, "X"), "#REF!+B1");
    }
}

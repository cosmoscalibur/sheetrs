//! Structural edits on references: what a formula becomes when rows or columns
//! are inserted or deleted, or a sheet is renamed or deleted, as the
//! spreadsheet applications do it.
//!
//! A reference to deleted cells becomes `Sheet!#REF!` (`#REF!` for a deleted sheet); a range
//! shrinks or grows around
//! the edit; absolute markers are irrelevant (both kinds move).

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

impl crate::parse::Formula {
    /// The formula text after `edit`, or `None` when it does not change.
    ///
    /// Only the references that change are replaced, so whitespace, letter case,
    /// function-name spelling and everything else in the source stay exactly as
    /// written.
    #[must_use]
    pub fn rewrite(&self, edit: &Edit<'_>, home_sheet: &str) -> Option<String> {
        self.rewrite_all(std::slice::from_ref(edit), home_sheet)
    }

    /// The formula text after every edit of `edits`, applied to each reference in order, or
    /// `None` when it does not change. The edits are written in the coordinates the formula has
    /// now: to apply several deletions of one sheet at once, list them from the last row (or
    /// column) to the first, so each leaves the positions of the next one as they were.
    ///
    ///
    /// A 3D reference (`Sheet1:Sheet3!A1`) is not understood: see [`crate::parse::Formula::has_three_d`], which
    /// a caller that cannot accept one asks first.
    #[must_use]
    pub fn rewrite_all(&self, edits: &[Edit<'_>], home_sheet: &str) -> Option<String> {
        let refs = self.expr.references();
        let mut out = String::with_capacity(self.source.len());
        let mut last = 0;
        let mut changed = false;
        for (r, span) in refs.into_iter().zip(&self.ref_spans) {
            let mut node = Expr::Ref(r.clone());
            let mut touched = false;
            for edit in edits {
                touched |= apply(&mut node, edit, home_sheet);
            }
            if !touched {
                continue;
            }
            out.push_str(&self.source[last..span.start]);
            out.push_str(&node.to_string());
            last = span.end;
            changed = true;
        }
        out.push_str(&self.source[last..]);
        changed.then_some(out)
    }

    /// Whether the formula holds a 3D reference, one that spans sheets (`Sheet1:Sheet3!A1`, or
    /// `'Sheet 1:Sheet 3'!A1`). Editing rows, columns or sheets under such a reference is not
    /// supported (issue #75): the edits here read it as a reference to a single sheet, which is
    /// wrong, so whoever applies them refuses a formula for which this is `true`.
    ///
    /// `Name:Sheet!A1` also reads as one: the AST of a range between a name and a sheet-qualified
    /// cell is the same as that of a span, and a name cannot be told from a sheet here.
    #[must_use]
    pub fn has_three_d(&self) -> bool {
        let mut found = false;
        self.expr.walk(&mut |e| match e {
            Expr::Binary(BinOp::Range, a, b)
                if matches!((&**a, &**b), (Expr::Name(_), Expr::Ref(r)) if r.sheet.is_some()) =>
            {
                found = true;
            }
            Expr::Ref(r) if r.sheet.as_deref().is_some_and(|sheet| sheet.contains(':')) => {
                found = true;
            }
            _ => {}
        });
        found
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
            parse_spanned(f).unwrap().rewrite_all(&edits, "S")
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

    #[test]
    fn a_span_of_sheets_is_a_3d_reference_and_a_single_sheet_is_not() {
        use crate::parse::parse_spanned;
        let three_d = |f: &str| parse_spanned(f).unwrap().has_three_d();
        assert!(three_d("=SUM(Alpha:Gamma!A3:A5)"));
        assert!(three_d("=SUM(Alpha:Gamma!A1)"));
        assert!(three_d("=1+SUM(Alpha:Gamma!A:A)"));
        assert!(three_d("=SUM('Alpha 1:Gamma 3'!A1)"));
        assert!(!three_d("=SUM(Alpha!A3:A5)"));
        assert!(!three_d("=SUM(Alpha!A3:Alpha!A5)"));
        assert!(!three_d("=A1:B2+Alpha!C3"));
        assert!(!three_d("=SUM(A1:A5)"));
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

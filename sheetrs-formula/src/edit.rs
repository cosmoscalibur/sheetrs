//! Structural edits on references: what a formula becomes when rows or columns
//! are inserted or deleted, or a sheet is renamed or deleted, as the
//! spreadsheet applications do it.
//!
//! A reference to deleted cells becomes `#REF!`; a range shrinks or grows around
//! the edit; absolute markers are irrelevant (both kinds move).

use crate::addr::{MAX_COLS, MAX_ROWS, same_name};
use crate::ast::{CellAddr, ColAddr, Expr, RefArea, Reference, RowAddr};

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
                changed |= replace(e, r, s, true);
            }
            Edit::InsertCols { sheet, at, count } | Edit::DeleteCols { sheet, at, count }
                if same_name(target, sheet) =>
            {
                let s = if matches!(edit, Edit::InsertCols { .. }) {
                    Shift::Insert { at, count }
                } else {
                    Shift::Delete { at, count }
                };
                changed |= replace(e, r, s, false);
            }
            _ => {}
        }
    });
    changed
}

fn replace(e: &mut Expr, r: Reference, s: Shift, rows: bool) -> bool {
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
            *e = Expr::Error("#REF!".into());
            true
        }
    }
}

// Keep the address types in the public documentation of this module.
#[allow(dead_code)]
type _Docs = (CellAddr, ColAddr, RowAddr);

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
        assert_eq!(del_rows("A1+A3+A4", 2, 1), "A1+#REF!+A3");
        assert_eq!(del_rows("SUM(A1:A5)", 2, 1), "SUM(A1:A4)");
        assert_eq!(del_rows("SUM(A3:A5)", 2, 1), "SUM(A3:A4)");
        assert_eq!(del_rows("SUM(A1:A3)", 2, 1), "SUM(A1:A2)");
        assert_eq!(del_rows("SUM(A3:A3)", 2, 1), "SUM(#REF!)");
        assert_eq!(del_rows("SUM(A2:A4)", 1, 3), "SUM(#REF!)");
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
        assert_eq!(cols("B1+C1+D1", 2, 1), "B1+#REF!+C1");
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
            "#REF!"
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
        assert_eq!(run("A5+Data!A5+'Data'!A1", e, "Other"), "A5+Data!A4+#REF!");
        // Unqualified reference on the edited sheet itself moves.
        assert_eq!(run("A5", e, "data"), "A4");
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

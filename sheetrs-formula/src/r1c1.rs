//! R1C1 spelling of a formula relative to its own cell: `=A1+B1` in `C1` and
//! `=A2+B2` in `C2` both become `RC[-2]+RC[-1]`. Formulas that do the same thing
//! down a column then compare equal as strings.

use crate::ast::{CellAddr, ColAddr, Expr, RefArea, Reference, RowAddr};

fn row_text(row: u32, abs: bool, origin: u32) -> String {
    if abs {
        format!("R{}", row + 1)
    } else {
        match i64::from(row) - i64::from(origin) {
            0 => "R".into(),
            d => format!("R[{d}]"),
        }
    }
}

fn col_text(col: u32, abs: bool, origin: u32) -> String {
    if abs {
        format!("C{}", col + 1)
    } else {
        match i64::from(col) - i64::from(origin) {
            0 => "C".into(),
            d => format!("C[{d}]"),
        }
    }
}

fn cell(c: &CellAddr, row: u32, col: u32) -> String {
    format!(
        "{}{}",
        row_text(c.row, c.row_abs, row),
        col_text(c.col, c.col_abs, col)
    )
}

fn area_text(r: &Reference, row: u32, col: u32) -> String {
    let body = match &r.area {
        RefArea::Cell(c) => cell(c, row, col),
        RefArea::Range(a, b) => format!("{}:{}", cell(a, row, col), cell(b, row, col)),
        RefArea::Cols(ColAddr { col: a, abs: x }, ColAddr { col: b, abs: y }) => {
            format!("{}:{}", col_text(*a, *x, col), col_text(*b, *y, col))
        }
        RefArea::Rows(RowAddr { row: a, abs: x }, RowAddr { row: b, abs: y }) => {
            format!("{}:{}", row_text(*a, *x, row), row_text(*b, *y, row))
        }
    };
    match &r.sheet {
        Some(s) => format!("{}!{body}", crate::print::sheet_text(s)),
        None => body,
    }
}

/// The formula in R1C1 form as seen from the zero-based cell `(row, col)`.
pub fn to_r1c1(expr: &Expr, row: u32, col: u32) -> String {
    let mut e = expr.clone();
    e.walk_mut(&mut |x| {
        if let Expr::Ref(r) = x {
            *x = Expr::Name(area_text(r, row, col));
        }
    });
    e.to_string()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parse::parse;

    fn r(f: &str, row: u32, col: u32) -> String {
        to_r1c1(&parse(f).unwrap(), row, col)
    }

    #[test]
    fn same_pattern_down_a_column() {
        assert_eq!(r("A1+B1", 0, 2), "RC[-2]+RC[-1]");
        assert_eq!(r("A2+B2", 1, 2), "RC[-2]+RC[-1]");
        assert_eq!(r("$A$1+B$1", 4, 2), "R1C1+R1C[-1]");
        assert_eq!(r("SUM(A1:A5)", 5, 1), "SUM(R[-5]C[-1]:R[-1]C[-1])");
        assert_eq!(r("Data!A1", 0, 0), "Data!R[0]C[0]".replace("[0]", ""));
        assert_eq!(r("A:A+1:2", 3, 3), "C[-3]:C[-3]+R[-3]:R[-2]");
    }
}

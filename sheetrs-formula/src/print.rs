//! Canonical printing: no whitespace, `,` separators, `$` markers and the
//! function name exactly as written.

use crate::addr::{col_name, needs_quotes};
use crate::ast::{CellAddr, Expr, RefArea, Reference, UnOp};
use std::fmt::{self, Display, Write};

/// A sheet name as it appears before `!`: quoted when needed.
pub fn sheet_text(name: &str) -> String {
    if needs_quotes(name) {
        format!("'{}'", name.replace('\'', "''"))
    } else {
        name.to_string()
    }
}

fn cell(f: &mut fmt::Formatter<'_>, c: &CellAddr) -> fmt::Result {
    write!(
        f,
        "{}{}{}{}",
        if c.col_abs { "$" } else { "" },
        col_name(c.col),
        if c.row_abs { "$" } else { "" },
        c.row + 1
    )
}

impl Display for Reference {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if let Some(s) = &self.sheet {
            write!(f, "{}!", sheet_text(s))?;
        }
        match &self.area {
            RefArea::Cell(c) => cell(f, c),
            RefArea::Range(a, b) => {
                cell(f, a)?;
                f.write_char(':')?;
                cell(f, b)
            }
            RefArea::Cols(a, b) => write!(
                f,
                "{}{}:{}{}",
                if a.abs { "$" } else { "" },
                col_name(a.col),
                if b.abs { "$" } else { "" },
                col_name(b.col)
            ),
            RefArea::Rows(a, b) => write!(
                f,
                "{}{}:{}{}",
                if a.abs { "$" } else { "" },
                a.row + 1,
                if b.abs { "$" } else { "" },
                b.row + 1
            ),
        }
    }
}

impl Display for Expr {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Number(n) => f.write_str(&n.text),
            Self::Text(t) => write!(f, "\"{}\"", t.replace('"', "\"\"")),
            Self::Bool(b) => f.write_str(if *b { "TRUE" } else { "FALSE" }),
            Self::Error(e) => f.write_str(e),
            Self::Array(rows) => {
                f.write_char('{')?;
                for (i, row) in rows.iter().enumerate() {
                    if i > 0 {
                        f.write_char(';')?;
                    }
                    for (j, e) in row.iter().enumerate() {
                        if j > 0 {
                            f.write_char(',')?;
                        }
                        write!(f, "{e}")?;
                    }
                }
                f.write_char('}')
            }
            Self::Ref(r) => write!(f, "{r}"),
            Self::Name(n) => f.write_str(n),
            Self::Call(c) => {
                write!(f, "{}(", c.name)?;
                for (i, a) in c.args.iter().enumerate() {
                    if i > 0 {
                        f.write_char(',')?;
                    }
                    if let Some(a) = a {
                        write!(f, "{a}")?;
                    }
                }
                f.write_char(')')
            }
            Self::Unary(op, e) => {
                f.write_char(match op {
                    UnOp::Neg => '-',
                    UnOp::Plus => '+',
                    UnOp::Implicit => '@',
                })?;
                write!(f, "{e}")
            }
            Self::Percent(e) => write!(f, "{e}%"),
            Self::Binary(op, a, b) => write!(f, "{a}{}{b}", op.symbol()),
            Self::Paren(e) => write!(f, "({e})"),
        }
    }
}

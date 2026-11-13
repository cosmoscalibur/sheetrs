//! Printing in `LibreOffice`'s own notation, as `FORMULA()` (`FORMULATEXT` in
//! Excel) shows it: `$Sheet.A1:B2` references (the sheet of the formula
//! itself is left out), `TRUE()`, canonical function names, numbers in the
//! reader's locale and the separators of that locale.

use crate::addr::col_name;
use crate::ast::{CellAddr, Expr, RefArea, Reference, UnOp};
use std::fmt::Write;

/// Separators a `LibreOffice` user has set; they follow the decimal separator
/// unless changed (`,` decimal: `;` between arguments, `.` between the columns
/// of an array constant).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CalcSyntax {
    pub decimal: char,
    pub function: char,
    pub column: char,
    pub row: char,
}

impl CalcSyntax {
    /// The defaults `LibreOffice` derives from the decimal separator.
    #[must_use]
    pub fn for_decimal(decimal: char) -> Self {
        if decimal == ',' {
            Self {
                decimal,
                function: ';',
                column: '.',
                row: ';',
            }
        } else {
            Self {
                decimal,
                function: ',',
                column: ',',
                row: ';',
            }
        }
    }
}

fn sheet_ref(name: &str) -> String {
    let plain = !name.is_empty()
        && !name.starts_with(|c: char| c.is_ascii_digit())
        && name.chars().all(|c| c.is_alphanumeric() || c == '_');
    if plain {
        format!("${name}.")
    } else {
        format!("$'{}'.", name.replace('\'', "''"))
    }
}

fn cell(out: &mut String, c: &CellAddr) {
    if c.col_abs {
        out.push('$');
    }
    let _ = write!(out, "{}", col_name(c.col));
    if c.row_abs {
        out.push('$');
    }
    let _ = write!(out, "{}", c.row + 1);
}

pub(crate) fn reference(out: &mut String, r: &Reference, home: &str) {
    if let Some(s) = &r.sheet
        && !crate::addr::same_name(s, home)
    {
        out.push_str(&sheet_ref(s));
    }
    match &r.area {
        RefArea::Cell(c) => cell(out, c),
        RefArea::Range(a, b) => {
            cell(out, a);
            out.push(':');
            cell(out, b);
        }
        RefArea::Cols(a, b) => {
            let _ = write!(
                out,
                "{}{}:{}{}",
                if a.abs { "$" } else { "" },
                col_name(a.col),
                if b.abs { "$" } else { "" },
                col_name(b.col)
            );
        }
        RefArea::Rows(a, b) => {
            let _ = write!(
                out,
                "{}{}:{}{}",
                if a.abs { "$" } else { "" },
                a.row + 1,
                if b.abs { "$" } else { "" },
                b.row + 1
            );
        }
    }
}

/// The function an ODS names by its add-in service (`COM.SUN.STAR.SHEET.ADDIN.PRICINGFUNCTIONS.
/// GETOPTBARRIER`), by the name an XLSX gives it (`OPT_BARRIER`); `name` is uppercase.
#[must_use]
pub fn addin_function(name: &str) -> Option<&'static str> {
    Some(
        match name.strip_prefix("COM.SUN.STAR.SHEET.ADDIN.PRICINGFUNCTIONS.GETOPT")? {
            "BARRIER" => "OPT_BARRIER",
            "TOUCH" => "OPT_TOUCH",
            "PROBHIT" => "OPT_PROB_HIT",
            "PROBINMONEY" => "OPT_PROB_INMONEY",
            _ => return None,
        },
    )
}

/// The prefixes that file formats put before a function name (`_xlfn.CONCAT`,
/// `COM.MICROSOFT.CONCAT`): the application shows the name without them.
const FUNCTION_NAMESPACES: [&str; 6] = [
    "_XLFN.",
    "_XLWS.",
    "COM.MICROSOFT.",
    "ORG.OPENOFFICE.",
    "ORG.LIBREOFFICE.",
    "LEGACY.",
];

/// The number of characters of the namespace prefix that `raw` carries, 0 when it has none.
#[must_use]
pub fn namespace_len(raw: &str) -> usize {
    let up = raw.to_ascii_uppercase();
    FUNCTION_NAMESPACES
        .iter()
        .find(|p| up.starts_with(**p))
        .map_or(0, |p| p.len())
}

/// The characters of the formula `source` as a user sees them: those of its text without the
/// namespace prefix of the functions, so `_xlfn.CONCAT(…)` and `COM.MICROSOFT.CONCAT(…)` count as
/// `CONCAT(…)`. A text that is not a formula counts as it is.
#[must_use]
pub fn visible_length(source: &str) -> usize {
    let total = source.chars().count();
    let Ok(expr) = crate::parse::parse(source) else {
        return total;
    };
    let mut prefixes = 0;
    expr.walk(&mut |e| {
        if let Expr::Call(c) = e {
            prefixes += namespace_len(&c.name);
        }
    });
    total - prefixes
}

/// The name the application shows for a function called `raw` in the file: the canonical name of one
/// it has, and, when `known` says it has no such function, the name as it was written (`foo()`).
pub(crate) fn function_name(raw: &str, known: &dyn Fn(&str) -> bool) -> String {
    let up = raw.to_ascii_uppercase();
    if let Some(n) = addin_function(&up) {
        return n.to_string();
    }
    let canonical = &up[namespace_len(raw)..];
    if known(canonical) {
        canonical.to_string()
    } else {
        raw.to_string()
    }
}

/// Whether the application has the function called `raw` in the file.
pub(crate) fn has_function(raw: &str, known: &dyn Fn(&str) -> bool) -> bool {
    let up = raw.to_ascii_uppercase();
    addin_function(&up).is_some() || known(&up[namespace_len(raw)..])
}

pub(crate) fn number(v: f64, syn: &CalcSyntax) -> String {
    let t = if v == v.trunc() && v.abs() < 1e15 {
        format!("{}", v as i64)
    } else {
        format!("{v}")
    };
    t.replace('.', &syn.decimal.to_string())
}

fn expr(out: &mut String, e: &Expr, home: &str, syn: &CalcSyntax, known: &dyn Fn(&str) -> bool) {
    match e {
        Expr::Number(n) => out.push_str(&number(n.value, syn)),
        Expr::Text(t) => {
            let _ = write!(out, "\"{}\"", t.replace('"', "\"\""));
        }
        Expr::Bool(b) => out.push_str(if *b { "TRUE()" } else { "FALSE()" }),
        Expr::Error(s) => out.push_str(s),
        Expr::Array(rows) => {
            out.push('{');
            for (i, row) in rows.iter().enumerate() {
                if i > 0 {
                    out.push(syn.row);
                }
                for (j, x) in row.iter().enumerate() {
                    if j > 0 {
                        out.push(syn.column);
                    }
                    expr(out, x, home, syn, known);
                }
            }
            out.push('}');
        }
        Expr::Ref(r) => reference(out, r, home),
        Expr::Name(n) => out.push_str(n),
        Expr::Call(c) => {
            out.push_str(&function_name(&c.name, known));
            out.push('(');
            for (i, a) in c.args.iter().enumerate() {
                if i > 0 {
                    out.push(syn.function);
                }
                if let Some(a) = a {
                    expr(out, a, home, syn, known);
                }
            }
            out.push(')');
        }
        Expr::Unary(op, x) => {
            out.push(match op {
                UnOp::Neg => '-',
                UnOp::Plus => '+',
                // Not LibreOffice syntax: the intersection is implicit there.
                UnOp::Implicit => return expr(out, x, home, syn, known),
            });
            expr(out, x, home, syn, known);
        }
        Expr::Percent(x) => {
            expr(out, x, home, syn, known);
            out.push('%');
        }
        Expr::Binary(op, a, b) => {
            expr(out, a, home, syn, known);
            out.push_str(op.symbol());
            expr(out, b, home, syn, known);
        }
        Expr::Paren(x) => {
            out.push('(');
            expr(out, x, home, syn, known);
            out.push(')');
        }
    }
}

/// `e` in `LibreOffice` notation, without the leading `=`. `home` is the sheet the
/// formula is in.
#[must_use]
pub fn to_calc(e: &Expr, home: &str, syn: &CalcSyntax) -> String {
    to_calc_with(e, home, syn, &|_| true)
}

/// [`to_calc`] with the functions the application has told apart from the ones it has not: `known`
/// says whether it has a function of this name (uppercase, without prefix); a function it does not
/// have is shown as it was written (`foo()`).
#[must_use]
pub fn to_calc_with(
    e: &Expr,
    home: &str,
    syn: &CalcSyntax,
    known: &dyn Fn(&str) -> bool,
) -> String {
    let mut out = String::new();
    expr(&mut out, e, home, syn, known);
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn t(src: &str, syn: CalcSyntax) -> String {
        to_calc(&crate::parse(src).unwrap(), "S", &syn)
    }

    #[test]
    fn prints_as_libreoffice_shows_formulas() {
        let en = CalcSyntax::for_decimal('.');
        let es = CalcSyntax::for_decimal(',');
        // Observed with LibreOffice 26.8 (FORMULA of each cell).
        assert_eq!(t("SUM(A1:A3,2)", en), "SUM(A1:A3,2)");
        assert_eq!(t("SUM(A1:A3,2)", es), "SUM(A1:A3;2)");
        assert_eq!(
            t("IF(A1>1,\"a\"\"b\",{1,2;3,4})", en),
            "IF(A1>1,\"a\"\"b\",{1,2;3,4})"
        );
        assert_eq!(
            t("IF(A1>1,\"a\",{1,2;3,4})", es),
            "IF(A1>1;\"a\";{1.2;3.4})"
        );
        assert_eq!(
            t("'My Sheet'!A1+F160!B2:C3+S!A1", en),
            "$'My Sheet'.A1+$F160.B2:C3+A1"
        );
        assert_eq!(t("$A$1+B$2", en), "$A$1+B$2");
        assert_eq!(
            t("VLOOKUP(1,A1:B2,2,FALSE)", en),
            "VLOOKUP(1,A1:B2,2,FALSE())"
        );
        assert_eq!(t("_xlfn.CONCAT(A1,B1)", en), "CONCAT(A1,B1)");
        assert_eq!(t("COM.MICROSOFT.CONCAT(A1,B1)", en), "CONCAT(A1,B1)");
        assert_eq!(visible_length("_xlfn.CONCAT(A1,B1)"), 13);
        assert_eq!(visible_length("COM.MICROSOFT.CONCAT(\"_xlfn.\")"), 16);
        assert_eq!(visible_length("not ( a formula"), 15);
        assert_eq!(t("1.50+1E3", en), "1.5+1000");
        assert_eq!(t("1.50", es), "1,5");
        assert_eq!(t("-A1^2+SUM(A1:A3)%", en), "-A1^2+SUM(A1:A3)%");
        assert_eq!(t("A:A+2:3", en), "A:A+2:3");
    }
}

//! Lowering of the formula AST (`sheetrs-formula`) to the IR: names and sheets
//! are resolved, function names get their namespace, literal ranges and
//! constants fold.
//!
//! Parsing the text is the job of `sheetrs-formula`, which keeps the formula
//! lossless for rewriting and linting; this pass is where the engine's own
//! decisions start: which sheet a reference means, which function a name is in
//! the file's dialect, what a defined name expands to.

use crate::addr::{self, CellRef, Rect};
use crate::functions;
use crate::ir::{BinOp, Expr, Ns, UnOp};
use crate::value::{Array, ErrorKind, Value};
use sheetrs_formula::ast;
use std::collections::HashMap;
use std::sync::Arc;

pub use sheetrs_formula::ParseError;

/// Name resolution context of a formula.
pub struct ParseCtx<'a> {
    pub sheets: &'a [String],
    pub current_sheet: u16,
    /// Defined names: name (any case) → formula text without `=`.
    pub names: &'a HashMap<String, String>,
    /// Namespace tried before `core` for unqualified function names.
    pub dialect: Ns,
}

impl ParseCtx<'_> {
    fn sheet_id(&self, name: &str) -> Option<u16> {
        self.sheets
            .iter()
            .position(|s| addr::same_name(s, name))
            .map(|i| i as u16)
    }
}

/// Parse a formula body (without the leading `=`) and lower it.
pub fn parse(src: &str, ctx: &ParseCtx, diags: &mut Vec<String>) -> Result<Expr, ParseError> {
    let ast = sheetrs_formula::parse(src)?;
    Ok(Lowering {
        ctx,
        diags,
        depth: 0,
    }
    .expr(&ast))
}

struct Lowering<'a, 'c> {
    ctx: &'a ParseCtx<'c>,
    diags: &'a mut Vec<String>,
    depth: u8,
}

fn bin_op(op: ast::BinOp) -> Option<BinOp> {
    Some(match op {
        ast::BinOp::Add => BinOp::Add,
        ast::BinOp::Sub => BinOp::Sub,
        ast::BinOp::Mul => BinOp::Mul,
        ast::BinOp::Div => BinOp::Div,
        ast::BinOp::Pow => BinOp::Pow,
        ast::BinOp::Concat => BinOp::Concat,
        ast::BinOp::Eq => BinOp::Eq,
        ast::BinOp::Ne => BinOp::Ne,
        ast::BinOp::Lt => BinOp::Lt,
        ast::BinOp::Le => BinOp::Le,
        ast::BinOp::Gt => BinOp::Gt,
        ast::BinOp::Ge => BinOp::Ge,
        ast::BinOp::Range => return None,
    })
}

fn constant(e: &ast::Expr) -> Option<Value> {
    Some(match e {
        ast::Expr::Number(n) => Value::Number(n.value),
        ast::Expr::Text(t) => Value::text(t),
        ast::Expr::Bool(b) => Value::Bool(*b),
        ast::Expr::Error(e) => Value::Error(ErrorKind::parse(e).unwrap_or(ErrorKind::Name)),
        ast::Expr::Unary(ast::UnOp::Neg, inner) => match constant(inner)? {
            Value::Number(n) => Value::Number(-n),
            _ => return None,
        },
        ast::Expr::Unary(ast::UnOp::Plus, inner) => match constant(inner)? {
            v @ Value::Number(_) => v,
            _ => return None,
        },
        _ => return None,
    })
}

impl Lowering<'_, '_> {
    fn expr(&mut self, e: &ast::Expr) -> Expr {
        match e {
            ast::Expr::Number(n) => Expr::Const(Value::Number(n.value)),
            ast::Expr::Text(t) => Expr::Const(Value::text(t)),
            ast::Expr::Bool(b) => Expr::Const(Value::Bool(*b)),
            ast::Expr::Error(e) => {
                Expr::Const(Value::Error(ErrorKind::parse(e).unwrap_or(ErrorKind::Name)))
            }
            ast::Expr::Array(rows) => {
                let cols = rows.first().map_or(0, Vec::len) as u32;
                let data: Vec<Value> = rows
                    .iter()
                    .flatten()
                    .map(|x| constant(x).unwrap_or(Value::Error(ErrorKind::Value)))
                    .collect();
                Expr::Const(Value::Array(Arc::new(Array::new(
                    rows.len() as u32,
                    cols,
                    data,
                ))))
            }
            ast::Expr::Ref(r) => self.reference(r),
            ast::Expr::Name(n) => self.name(n),
            ast::Expr::Call(c) => self.call(c),
            ast::Expr::Unary(ast::UnOp::Implicit, inner) | ast::Expr::Paren(inner) => {
                self.expr(inner)
            }
            // A negative literal is one constant, not a negation.
            ast::Expr::Unary(ast::UnOp::Neg, inner) if matches!(**inner, ast::Expr::Number(_)) => {
                Expr::Const(constant(e).unwrap_or_default())
            }
            ast::Expr::Unary(op, inner) => {
                let op = if *op == ast::UnOp::Neg {
                    UnOp::Neg
                } else {
                    UnOp::Plus
                };
                Expr::Unary(op, Box::new(self.expr(inner)))
            }
            ast::Expr::Percent(inner) => Expr::Unary(UnOp::Percent, Box::new(self.expr(inner))),
            ast::Expr::Binary(op, a, b) => {
                let (l, r) = (self.expr(a), self.expr(b));
                match bin_op(*op) {
                    Some(op) => Expr::Binary(op, Box::new(l), Box::new(r)),
                    None => range_of(l, r),
                }
            }
        }
    }

    fn reference(&mut self, r: &ast::Reference) -> Expr {
        let sheet = match &r.sheet {
            None => self.ctx.current_sheet,
            Some(name) => match self.ctx.sheet_id(name) {
                Some(s) => s,
                None => {
                    // Unknown sheet or an external workbook: keep going with #REF!.
                    self.diags.push(format!("unresolved sheet `{name}`"));
                    return Expr::Const(Value::Error(ErrorKind::Ref));
                }
            },
        };
        let rect = match r.area {
            ast::RefArea::Cell(c) => {
                return Expr::Cell(CellRef::new(sheet, c.row, c.col));
            }
            ast::RefArea::Range(a, b) => Rect::new(sheet, a.row, a.col, b.row, b.col),
            ast::RefArea::Cols(a, b) => Rect::new(sheet, 0, a.col, addr::MAX_ROWS - 1, b.col),
            ast::RefArea::Rows(a, b) => Rect::new(sheet, a.row, 0, b.row, addr::MAX_COLS - 1),
        };
        if rect.is_single() {
            Expr::Cell(rect.top_left())
        } else {
            Expr::Range(rect)
        }
    }

    fn name(&mut self, text: &str) -> Expr {
        let found = self
            .ctx
            .names
            .iter()
            .find(|(k, _)| addr::same_name(k, text))
            .map(|(_, v)| v.clone());
        match found {
            Some(def) if self.depth < 8 => match sheetrs_formula::parse(&def) {
                Ok(ast) => {
                    self.depth += 1;
                    let e = self.expr(&ast);
                    self.depth -= 1;
                    e
                }
                Err(e) => {
                    self.diags.push(format!("defined name `{text}`: {e}"));
                    Expr::Const(Value::Error(ErrorKind::Name))
                }
            },
            _ => {
                self.diags.push(format!("unresolved name `{text}`"));
                Expr::Const(Value::Error(ErrorKind::Name))
            }
        }
    }

    fn call(&mut self, c: &ast::Call) -> Expr {
        let args: Vec<Expr> = c
            .args
            .iter()
            .map(|a| a.as_ref().map_or(Expr::Missing, |a| self.expr(a)))
            .collect();
        let (ns, name) = split_function_name(&c.name, self.ctx.dialect);
        match functions::lookup(ns, &name) {
            Some(id) => {
                let def = id.def();
                if args.len() < def.min as usize || args.len() > def.max as usize {
                    self.diags.push(format!(
                        "{}: wrong number of arguments ({})",
                        id.qualified(),
                        args.len()
                    ));
                }
                Expr::Call(id, args)
            }
            None => {
                self.diags.push(format!("unsupported function `{name}`"));
                Expr::Unknown(name.into(), args)
            }
        }
    }
}

/// Normalize a function name and pick the namespace its prefix implies.
///
/// `_xlfn.`/`_xlws.` mark Excel extension functions in XLSX;
/// `COM.MICROSOFT.`, `ORG.OPENOFFICE.` and `ORG.LIBREOFFICE.` are the ODS
/// (OpenFormula) prefixes for the same idea.
pub fn split_function_name(raw: &str, dialect: Ns) -> (Ns, String) {
    let up = raw.to_ascii_uppercase();
    for (prefix, ns) in [
        ("_XLFN.", Ns::Xlsx),
        ("_XLWS.", Ns::Xlsx),
        ("COM.MICROSOFT.", Ns::Ods),
        ("ORG.OPENOFFICE.", Ns::Ods),
        ("ORG.LIBREOFFICE.", Ns::Ods),
    ] {
        if let Some(rest) = up.strip_prefix(prefix) {
            return (
                if dialect == Ns::Core { ns } else { dialect },
                rest.to_string(),
            );
        }
    }
    (dialect, up)
}

/// Fold `a:b` over literal references, otherwise defer to the dynamic
/// `_RANGE` function.
fn range_of(a: Expr, b: Expr) -> Expr {
    let rect = |e: &Expr| match e {
        Expr::Cell(c) => Some(Rect::single(*c)),
        Expr::Range(r) => Some(*r),
        _ => None,
    };
    if let (Some(x), Some(y)) = (rect(&a), rect(&b))
        && x.sheet == y.sheet
    {
        return Expr::Range(Rect::new(
            x.sheet,
            x.r0.min(y.r0),
            x.c0.min(y.c0),
            x.r1.max(y.r1),
            x.c1.max(y.c1),
        ));
    }
    match functions::lookup(Ns::Core, "_RANGE") {
        Some(id) => Expr::Call(id, vec![a, b]),
        None => Expr::Const(Value::Error(ErrorKind::Ref)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ctx<'a>(sheets: &'a [String], names: &'a HashMap<String, String>) -> ParseCtx<'a> {
        ParseCtx {
            sheets,
            current_sheet: 0,
            names,
            dialect: Ns::Xlsx,
        }
    }

    fn p(src: &str) -> Expr {
        let sheets = vec!["Sheet1".to_string(), "My Sheet".to_string()];
        let names = HashMap::from([("Rate".to_string(), "Sheet1!$B$1".to_string())]);
        parse(src, &ctx(&sheets, &names), &mut vec![]).unwrap()
    }

    fn n(v: f64) -> Expr {
        Expr::Const(Value::Number(v))
    }

    #[test]
    fn precedence_and_unary() {
        // -2^2 = 4: unary minus binds tighter than ^.
        let Expr::Binary(BinOp::Pow, l, _) = p("-2^2") else {
            panic!()
        };
        assert_eq!(*l, n(-2.0));
        let Expr::Binary(BinOp::Add, _, r) = p("1+2*3") else {
            panic!()
        };
        assert!(matches!(*r, Expr::Binary(BinOp::Mul, ..)));
        assert!(matches!(p("1&2=3"), Expr::Binary(BinOp::Eq, ..)));
        assert!(matches!(p("50%"), Expr::Unary(UnOp::Percent, _)));
    }

    #[test]
    fn references() {
        assert_eq!(p("$B$2"), Expr::Cell(CellRef::new(0, 1, 1)));
        assert_eq!(p("'My Sheet'!A1:B2"), Expr::Range(Rect::new(1, 0, 0, 1, 1)));
        assert_eq!(p("Sheet1!C3"), Expr::Cell(CellRef::new(0, 2, 2)));
        assert!(matches!(p("A:A"), Expr::Range(r) if r.rows() == addr::MAX_ROWS));
        assert!(matches!(p("1:2"), Expr::Range(r) if r.cols() == addr::MAX_COLS));
        assert_eq!(p("Rate"), Expr::Cell(CellRef::new(0, 0, 1)));
    }

    #[test]
    fn calls_and_prefixes() {
        let Expr::Call(id, args) = p("SUM(A1:A3, 1)") else {
            panic!()
        };
        assert_eq!(id.qualified(), "core::SUM");
        assert_eq!(args.len(), 2);
        let Expr::Call(id, _) = p("_xlfn.CONCAT(\"a\")") else {
            panic!()
        };
        assert_eq!(id.qualified(), "core::CONCAT");
        assert!(matches!(p("IF(A1,,2)"), Expr::Call(_, a) if a[1] == Expr::Missing));
        assert!(matches!(p("NOPE(1)"), Expr::Unknown(..)));
    }

    #[test]
    fn literals() {
        assert_eq!(p("\"a\"\"b\""), Expr::Const(Value::text("a\"b")));
        assert_eq!(p("#N/A"), Expr::Const(Value::Error(ErrorKind::Na)));
        assert_eq!(p("TRUE"), Expr::Const(Value::Bool(true)));
        let Expr::Const(Value::Array(a)) = p("{1,2;3,4}") else {
            panic!()
        };
        assert_eq!((a.rows, a.cols), (2, 2));
        assert_eq!(p("1.5E3"), n(1500.0));
    }

    #[test]
    fn dynamic_range_operator() {
        assert!(matches!(p("A1:INDEX(B1:B5,2)"), Expr::Call(..)));
    }
}

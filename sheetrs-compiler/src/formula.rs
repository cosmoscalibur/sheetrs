//! Lowering of the formula AST (`sheetrs-formula`) to the IR: names and sheets
//! are resolved, function names get their namespace, literal ranges and
//! constants fold.
//!
//! Parsing the text is the job of `sheetrs-formula`, which keeps the formula
//! lossless for rewriting and linting; this pass is where the engine's own
//! decisions start: which sheet a reference means, which function a name is in
//! what a defined name expands to.

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
    /// The workbook's cells, for the functions that ask about a cell's formula
    /// (`ISFORMULA`, `FORMULATEXT`): they are answered while lowering.
    pub cells: Option<&'a crate::cellstore::CellStore>,
    /// The formulas were read from an XLSX.
    pub xlsx: bool,
    /// The formulas were read from an ODS: a name without a namespace is LibreOffice's own
    /// (`CEILING` and `FLOOR` are not Excel's, which an ODS names `COM.MICROSOFT.CEILING`).
    pub ods: bool,
    /// Notation of the text `FORMULATEXT` returns.
    pub calc: sheetrs_formula::calc_print::CalcSyntax,
    /// `_xlfn._LONGTEXT(…)` of constant text is one literal (feature `LONGTEXT`).
    pub longtext: bool,
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
        env: Vec::new(),
    }
    .expr(&ast))
}

struct Lowering<'a, 'c> {
    ctx: &'a ParseCtx<'c>,
    diags: &'a mut Vec<String>,
    depth: u8,
    /// Names bound by `LET`, innermost last.
    env: Vec<(String, Expr)>,
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
            ast::Expr::Error(e) => Expr::Const(Value::Error(self.error_literal(e))),
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

    /// An error written in a formula. LibreOffice knows the spellings in capitals
    /// only: `#ref!` is a name it cannot resolve; an XLSX import does not read a bare
    /// `#REF!` either (the other error literals are fine).
    fn error_literal(&self, text: &str) -> ErrorKind {
        // `Sheet!#REF!`, what a reference that went leaves in an XLSX: `#REF!` when the sheet is
        // there, as LibreOffice reads it, and an unknown name when it is not.
        if let Some(qualifier) = text.strip_suffix("#REF!").filter(|q| q.ends_with('!')) {
            // The qualifier, read as the one of a cell, gives the sheet name unquoted.
            let (sheet, _) = sheetrs_formula::addr::split_sheet(&format!("{qualifier}A1"));
            return if sheet.is_some_and(|name| self.ctx.sheet_id(&name).is_some()) {
                ErrorKind::Ref
            } else {
                ErrorKind::Name
            };
        }
        if text != text.to_ascii_uppercase() || (self.ctx.xlsx && text == "#REF!") {
            return ErrorKind::Name;
        }
        ErrorKind::parse(text).unwrap_or(ErrorKind::Name)
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
        let bare = text.strip_prefix("_xlpm.").unwrap_or(text);
        if let Some((_, e)) = self
            .env
            .iter()
            .rev()
            .find(|(n, _)| n.eq_ignore_ascii_case(bare))
        {
            return e.clone();
        }
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

    /// `ISFORMULA(ref)` and `FORMULATEXT(ref)` of a literal reference: what a cell
    /// holds is known when the workbook is read, and neither function depends on
    /// the cell's value.
    fn formula_info(&mut self, c: &ast::Call, name: &str) -> Option<Expr> {
        let cells = self.ctx.cells?;
        if !matches!(name, "ISFORMULA" | "FORMULATEXT" | "FORMULA") || c.args.len() != 1 {
            return None;
        }
        let mut arg = c.args[0].as_ref()?;
        while let ast::Expr::Paren(x) | ast::Expr::Unary(ast::UnOp::Implicit, x) = arg {
            arg = x;
        }
        let ast::Expr::Ref(r) = arg else { return None };
        let ast::RefArea::Cell(a) = r.area else {
            return None;
        };
        let sheet = match &r.sheet {
            None => self.ctx.current_sheet,
            Some(n) => self.ctx.sheet_id(n)?,
        };
        let cell = cells.get(&CellRef::new(sheet, a.row, a.col));
        let formula = cell.and_then(|c| c.formula.as_deref().map(|f| (f, c.array)));
        Some(Expr::Const(if name == "ISFORMULA" {
            Value::Bool(formula.is_some())
        } else {
            match formula {
                None => Value::Error(ErrorKind::Na),
                Some((f, array)) => match sheetrs_formula::parse(f) {
                    Ok(ast) => {
                        let text = sheetrs_formula::calc_print::to_calc(
                            &ast,
                            &self.ctx.sheets[sheet as usize],
                            &self.ctx.calc,
                        );
                        Value::text(if array {
                            format!("{{={text}}}")
                        } else {
                            format!("={text}")
                        })
                    }
                    Err(_) => Value::Error(ErrorKind::Value),
                },
            }
        }))
    }

    /// `LET(name1, value1, …, result)`: the names stand for their (lowered) values
    /// in what follows, so the IR needs no binding construct.
    fn let_call(&mut self, c: &ast::Call) -> Expr {
        let args: Vec<&ast::Expr> = c.args.iter().filter_map(Option::as_ref).collect();
        if args.len() < 3 || args.len().is_multiple_of(2) || args.len() != c.args.len() {
            self.diags.push("LET: wrong number of arguments".into());
            return Expr::Const(Value::Error(ErrorKind::Value));
        }
        let mark = self.env.len();
        for pair in args[..args.len() - 1].chunks(2) {
            let ast::Expr::Name(n) = pair[0] else {
                self.env.truncate(mark);
                return Expr::Const(Value::Error(ErrorKind::Value));
            };
            let value = self.expr(pair[1]);
            let name = n.strip_prefix("_xlpm.").unwrap_or(n);
            // LibreOffice does not rebind a name an enclosing LET already holds
            // (Excel shadows it): `LET(x,1,LET(x,x+10,x*2))` is 2.
            if !self.env.iter().any(|(k, _)| k.eq_ignore_ascii_case(name)) {
                self.env.push((name.to_string(), value));
            }
        }
        let body = self.expr(args[args.len() - 1]);
        self.env.truncate(mark);
        body
    }

    fn call(&mut self, c: &ast::Call) -> Expr {
        let name = split_function_name(&c.name);
        if name == "LET" {
            return self.let_call(c);
        }
        if let Some(e) = self.formula_info(c, &name) {
            return e;
        }
        let args: Vec<Expr> = c
            .args
            .iter()
            .map(|a| a.as_ref().map_or(Expr::Missing, |a| self.expr(a)))
            .collect();
        if name == "_LONGTEXT"
            && self.ctx.longtext
            && let Some(joined) = join_longtext(&args)
        {
            return Expr::Const(Value::text(joined));
        }
        let (ns, name) = self.function_namespace(&c.name);
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
                if name == "_LONGTEXT" && self.ctx.longtext {
                    self.diags.push(
                        "`_xlfn._longtext(…)` with an argument that is not a constant text cannot be joined (feature LONGTEXT); LibreOffice 26.8 reads it as #NAME? (as this does); inspect the cell"
                            .to_string(),
                    );
                } else if name == "_LONGTEXT" {
                    self.diags.push(
                        "`_xlfn._longtext(…)` is not an Excel function: LibreOffice 26.2 wrote a text literal longer than 255 characters as pieces joined by it, and LibreOffice 26.8 reads it as #NAME? (as this does); the LONGTEXT feature (enable_features) joins constant pieces; inspect the cell"
                            .to_string(),
                    );
                } else {
                    self.diags.push(format!("unsupported function `{name}`"));
                }
                Expr::Unknown(Box::new((name.into(), args)))
            }
        }
    }
}

/// The text of the arguments of `_LONGTEXT(…)` joined, when every one is a constant text.
fn join_longtext(args: &[Expr]) -> Option<String> {
    let mut joined = String::new();
    for a in args {
        match a {
            Expr::Const(Value::Text(t)) => joined.push_str(t),
            _ => return None,
        }
    }
    (!args.is_empty()).then_some(joined)
}

/// A function name as the formula wrote it, uppercase and without the prefixes: `_xlfn.`/`_xlws.`
/// (XLSX: the function is newer than the file format) and the namespaces of OpenDocument.
pub fn split_function_name(raw: &str) -> String {
    let up = raw.to_ascii_uppercase();
    let (rest, _) = strip_future(&up);
    if let Some(n) = sheetrs_formula::calc_print::addin_function(rest) {
        return n.to_string();
    }
    Ns::of_prefixed(rest).map_or(rest, |(_, n)| n).to_string()
}

/// The name without the `_xlfn.` and `_xlws.` prefixes (`_xlfn._xlws.FILTER` has both), and
/// whether it had any.
fn strip_future(up: &str) -> (&str, bool) {
    let mut rest = up;
    while let Some(r) = ["_XLFN.", "_XLWS."]
        .iter()
        .find_map(|p| rest.strip_prefix(p))
    {
        rest = r;
    }
    (rest, rest.len() != up.len())
}

impl Lowering<'_, '_> {
    /// The namespace and bare name a call means. A prefix names its namespace (`COM.MICROSOFT.`
    /// in an ODS, `_xlfn.COM.MICROSOFT.` or just `_xlfn.` in an XLSX, where the function's own
    /// namespace is the one the catalogue gives it). A name without prefix is `core`, except in an
    /// XLSX (and a model built by hand) where a few bare names are Excel's or the legacy version
    /// of a function: `CEILING`, `FLOOR`, `FDIST`… (`docs/FUNCTION_NAMES.md`).
    fn function_namespace(&self, raw: &str) -> (Ns, String) {
        let up = raw.to_ascii_uppercase();
        let (rest, future) = strip_future(&up);
        if let Some(n) = sheetrs_formula::calc_print::addin_function(rest) {
            return (Ns::Core, n.to_string());
        }
        if let Some((ns, name)) = Ns::of_prefixed(rest) {
            return (ns, name.to_string());
        }
        let ns = if !self.ctx.ods && !future {
            functions::xlsx_bare(rest).unwrap_or(Ns::Core)
        } else {
            Ns::Core
        };
        (ns, rest.to_string())
    }
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
            cells: None,
            xlsx: false,
            ods: false,
            calc: sheetrs_formula::calc_print::CalcSyntax::for_decimal('.'),
            longtext: false,
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
        assert_eq!(id.qualified(), "microsoft::CONCAT");
        assert!(matches!(p("IF(A1,,2)"), Expr::Call(_, a) if a[1] == Expr::Missing));
        assert!(matches!(p("NOPE(1)"), Expr::Unknown(..)));
        // In an ODS a bare CEILING is the standard (core) one; with the prefix it is Microsoft's.
        let (sheets, names) = (vec!["S".to_string()], HashMap::new());
        let name = |f: &str, ods: bool| {
            let mut c = ctx(&sheets, &names);
            c.ods = ods;
            match parse(f, &c, &mut Vec::new()).unwrap() {
                Expr::Call(id, _) => id.qualified(),
                other => panic!("{other:?}"),
            }
        };
        assert_eq!(name("CEILING(1,2)", true), "core::CEILING");
        assert_eq!(
            name("COM.MICROSOFT.CEILING(1,2)", true),
            "microsoft::CEILING"
        );
        assert_eq!(
            name("_xlfn.CEILING.MATH(1,2)", true),
            "microsoft::CEILING.MATH"
        );
        // In an XLSX, and in a model built by hand, a bare CEILING is Excel's.
        assert_eq!(name("FLOOR(1,2)", false), "microsoft::FLOOR");
        assert_eq!(name("SUM(1)", false), "core::SUM");
        // A name that is newer than the format gets the namespace the catalogue gives it.
        assert_eq!(name("_xlfn.IFS(1,2)", false), "microsoft::IFS");
        assert_eq!(name("COM.MICROSOFT.IFS(1,2)", true), "microsoft::IFS");
        // An add-in function an ODS names by its service is the one an XLSX gives by its own name.
        assert_eq!(
            name(
                "COM.SUN.STAR.SHEET.ADDIN.PRICINGFUNCTIONS.GETOPTPROBHIT(1,1,0,1,1,0)",
                true
            ),
            "core::OPT_PROB_HIT"
        );
        // An ODS writes EASTERSUNDAY bare; an XLSX with the namespace of the extension. Both are it.
        assert_eq!(name("EASTERSUNDAY(2024)", true), "core::EASTERSUNDAY");
        assert_eq!(
            name("_xlfn.ORG.OPENOFFICE.EASTERSUNDAY(2024)", false),
            "core::EASTERSUNDAY"
        );
    }

    #[test]
    fn a_reference_that_went_is_a_ref_error_on_a_sheet_that_exists() {
        let sheets = vec!["Sheet1".to_string(), "My Data".to_string()];
        let names = HashMap::new();
        let mut diags = vec![];
        let mut lower = |src: &str| {
            let e = parse(src, &ctx(&sheets, &names), &mut diags).unwrap();
            match e {
                Expr::Const(Value::Error(kind)) => kind,
                other => panic!("{other:?}"),
            }
        };
        assert_eq!(lower("Sheet1!#REF!"), ErrorKind::Ref);
        assert_eq!(lower("'My Data'!#REF!"), ErrorKind::Ref);
        assert_eq!(lower("sheet1!#REF!"), ErrorKind::Ref);
        // A sheet the workbook does not have is an unknown name, as in LibreOffice.
        assert_eq!(lower("Gone!#REF!"), ErrorKind::Name);
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

    #[test]
    fn longtext_pieces_are_joined_only_when_constant() {
        let sheets = vec!["Sheet1".to_string()];
        let names = HashMap::new();
        let src = "_xlfn._LONGTEXT(\"ab\",\"cd\")";
        let mut on = ctx(&sheets, &names);
        on.longtext = true;
        let mut diags = vec![];
        let e = parse(src, &on, &mut diags).unwrap();
        assert_eq!(e, Expr::Const(Value::text("abcd")));
        assert!(diags.is_empty());

        // A piece that is not a constant text: not joined, reported.
        let mut diags = vec![];
        let e = parse("_xlfn._LONGTEXT(\"ab\",A1)", &on, &mut diags).unwrap();
        assert!(matches!(e, Expr::Unknown(..)));
        assert_eq!(diags.len(), 1);

        // Feature off: the function is unknown, as in LibreOffice 26.8.
        let mut diags = vec![];
        let e = parse(src, &ctx(&sheets, &names), &mut diags).unwrap();
        assert!(matches!(e, Expr::Unknown(..)));
        assert!(diags[0].contains("LONGTEXT"));
    }
}

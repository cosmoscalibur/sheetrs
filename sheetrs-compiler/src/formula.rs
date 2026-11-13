//! Lowering of the formula AST (`sheetrs-formula`) to the IR: names and sheets
//! are resolved, function names get their namespace, literal ranges and
//! constants fold.
//!
//! Parsing the text is the job of `sheetrs-formula`, which keeps the formula
//! lossless for rewriting and linting; this pass is where the engine's own
//! decisions start: which sheet a reference means, which function a name is in
//! what a defined name expands to.

use sheetrs_engine::addr::{self, CellRef, Rect};
use sheetrs_engine::functions;
use sheetrs_engine::ir::{BinOp, Expr, Ns, UnOp};
use sheetrs_engine::value::{Array, ErrorKind, Value};
use sheetrs_formula::ast;
use std::collections::HashMap;
use std::sync::Arc;

pub use sheetrs_formula::ParseError;

/// Name resolution context of a formula.
#[derive(Debug)]
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
    /// The formulas were read from an ODS: a name without a namespace is `LibreOffice`'s own
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
///
/// # Errors
///
/// When `src` is not a formula the parser reads.
pub fn parse(src: &str, ctx: &ParseCtx, diags: &mut Vec<String>) -> Result<Expr, ParseError> {
    let ast = match sheetrs_formula::parse(src) {
        Ok(ast) => ast,
        // A text with an array that the parser refuses is what `LibreOffice` reads with what is
        // left open closed, or the error it gives the cell (`{1,2}{3}` is `Err:509`, `{1 2}` is
        // `Err:539`), not a formula the engine cannot take.
        Err(refused) => match sheetrs_formula::lax::lax(src) {
            Some(sheetrs_formula::lax::Lax::Error(code)) => {
                return Ok(Expr::Const(Value::Error(ErrorKind::Libre(code))));
            }
            Some(sheetrs_formula::lax::Lax::Closed(text)) => {
                sheetrs_formula::parse(&text).map_err(|_| refused)?
            }
            None => return Err(refused),
        },
    };
    let mut lowering = Lowering {
        ctx,
        diags,
        depth: 0,
        env: Vec::new(),
        bad_array: None,
    };
    let expr = lowering.expr(&ast);
    // `LibreOffice` refuses an array written in the formula that holds an error (`{1;#N/A}`) or
    // anything but constants, or whose rows differ in width, when it compiles the formula, whatever
    // the function around it and whether it is ever reached (`IF(FALSE, {1;#N/A}, 3)`,
    // `ISERROR({#N/A})`): the whole formula is `Err:539`.
    Ok(match lowering.bad_array {
        Some(code) => Expr::Const(Value::Error(ErrorKind::Libre(code))),
        None => expr,
    })
}

/// The database functions (`DSUM`…): see [`Lowering::call`].
const DATABASE_FUNCTIONS: &[&str] = &[
    "DSUM", "DCOUNT", "DCOUNTA", "DAVERAGE", "DGET", "DMAX", "DMIN", "DPRODUCT", "DSTDEV",
    "DSTDEVP", "DVAR", "DVARP",
];

struct Lowering<'a, 'c> {
    ctx: &'a ParseCtx<'c>,
    diags: &'a mut Vec<String>,
    depth: u8,
    /// Names bound by `LET`, innermost last.
    env: Vec<(String, Expr)>,
    /// An array written in the formula that `LibreOffice` refuses: see [`parse`].
    bad_array: Option<u16>,
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

/// An element of an array written in the formula: a number (with one sign at most), a text or a
/// boolean; `None` for anything else, and an error for an error.
fn array_element(e: &ast::Expr) -> Option<Value> {
    match e {
        // An element left out.
        ast::Expr::Error(text) if text.is_empty() => Some(Value::Empty),
        ast::Expr::Unary(ast::UnOp::Neg | ast::UnOp::Plus, inner)
            if !matches!(**inner, ast::Expr::Number(_)) =>
        {
            None
        }
        _ => constant(e),
    }
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
                // An array inside an array is `Err:533`; the rest that `LibreOffice` refuses is
                // `Err:539`.
                if rows
                    .iter()
                    .flatten()
                    .any(|e| matches!(e, ast::Expr::Array(_)))
                {
                    self.bad_array = Some(533);
                    return Expr::Const(Value::Error(ErrorKind::Libre(533)));
                }
                if rows.iter().any(|r| r.len() as u32 != cols) {
                    self.bad_array.get_or_insert(539);
                    return Expr::Const(Value::Error(ErrorKind::Libre(539)));
                }
                // A boolean in an array is the number 1 or 0 (`{TRUE,FALSE}` shows 1 and 0).
                let elements: Vec<Option<Value>> = rows
                    .iter()
                    .flatten()
                    .map(|e| match array_element(e) {
                        Some(Value::Bool(b)) => Some(Value::Number(f64::from(u8::from(b)))),
                        v => v,
                    })
                    .collect();
                // `LibreOffice` takes only constants in an array written in the formula, in rows of
                // one width: anything else (`{1;A1}`, `{1,2+3}`, `{1,2;3}`) or an error inside is
                // `Err:539` for the whole formula, whether or not the array is ever reached.
                if elements
                    .iter()
                    .any(|v| !matches!(v, Some(v) if !matches!(v, Value::Error(_))))
                {
                    self.bad_array.get_or_insert(539);
                }
                let data = elements
                    .into_iter()
                    .map(Option::unwrap_or_default)
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
                    Some(op) => Expr::Binary(op, Box::new([l, r])),
                    None => range_of(l, r),
                }
            }
        }
    }

    /// An error written in a formula. `LibreOffice` knows the spellings in capitals
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

    /// A reference over several sheets (`Data:Wide!A1:A3`, `Data!A1:Wide!A3`) as the area of each sheet
    /// from the first to the last in the order of the workbook, `None` if `e` is not one (see
    /// `functions::threed`).
    fn three_d(&self, e: &ast::Expr) -> Option<Vec<Expr>> {
        let area = |a: ast::RefArea| match a {
            ast::RefArea::Cell(c) => Some((c.row, c.col, c.row, c.col)),
            ast::RefArea::Range(a, b) => Some((a.row, a.col, b.row, b.col)),
            // Whole columns and rows over the sheets are how a file keeps a range of every row.
            ast::RefArea::Cols(a, b) => Some((0, a.col, addr::MAX_ROWS - 1, b.col)),
            ast::RefArea::Rows(a, b) => Some((a.row, 0, b.row, addr::MAX_COLS - 1)),
        };
        let (first, last, a, b) = match e {
            ast::Expr::Ref(r) => {
                let (from, to) = r.sheet.as_deref()?.split_once(':')?;
                let a = area(r.area)?;
                (self.ctx.sheet_id(from)?, self.ctx.sheet_id(to)?, a, a)
            }
            ast::Expr::Binary(ast::BinOp::Range, x, y) => {
                // `Data!A2:Wide!A`: a column with no row is the column on the row of the first corner.
                if let (ast::Expr::Ref(x), ast::Expr::Name(n)) = (&**x, &**y)
                    && let ast::RefArea::Cell(c) = x.area
                    && let (Some(sheet), tail) = sheetrs_formula::addr::split_sheet(n)
                    && let Some((col, abs)) = sheetrs_formula::addr::parse_col(tail)
                    && let (Some(from), Some(to)) = (
                        self.ctx.sheet_id(x.sheet.as_deref()?),
                        self.ctx.sheet_id(&sheet),
                    )
                    && from != to
                {
                    let y = ast::Reference {
                        sheet: Some(sheet),
                        area: ast::RefArea::Cell(ast::CellAddr {
                            row: c.row,
                            col,
                            row_abs: c.row_abs,
                            col_abs: abs,
                        }),
                    };
                    return self.three_d(&ast::Expr::Binary(
                        ast::BinOp::Range,
                        Box::new(ast::Expr::Ref(x.clone())),
                        Box::new(ast::Expr::Ref(y)),
                    ));
                }
                let (ast::Expr::Ref(x), ast::Expr::Ref(y)) = (&**x, &**y) else {
                    return None;
                };
                let (from, to) = (
                    self.ctx.sheet_id(x.sheet.as_deref()?)?,
                    self.ctx.sheet_id(y.sheet.as_deref()?)?,
                );
                if from == to {
                    return None;
                }
                (from, to, area(x.area)?, area(y.area)?)
            }
            _ => return None,
        };
        let (r0, c0) = (
            a.0.min(b.0).min(a.2).min(b.2),
            a.1.min(b.1).min(a.3).min(b.3),
        );
        let (r1, c1) = (
            a.0.max(b.0).max(a.2).max(b.2),
            a.1.max(b.1).max(a.3).max(b.3),
        );
        Some(
            (first.min(last)..=first.max(last))
                .map(|s| {
                    let rect = Rect::new(s, r0, c0, r1, c1);
                    if rect.is_single() {
                        Expr::Cell(rect.top_left())
                    } else {
                        Expr::Range(rect)
                    }
                })
                .collect(),
        )
    }

    fn reference(&mut self, r: &ast::Reference) -> Expr {
        // A reference over several sheets where a function does not take it is not a value.
        if r.sheet.as_deref().is_some_and(|s| s.contains(':')) {
            return Expr::Const(Value::Error(ErrorKind::Value));
        }
        let sheet = match &r.sheet {
            None => self.ctx.current_sheet,
            Some(name) => {
                if let Some(s) = self.ctx.sheet_id(name) {
                    s
                } else {
                    // An XLSX reads a sheet that is not there as a name that is not defined, as
                    // `LibreOffice` does (`Zzz!A1` is `#NAME?`); an external workbook and the sheets of an
                    // ODS keep going with `#REF!`.
                    self.diags.push(format!("unresolved sheet `{name}`"));
                    let external = name.starts_with('[');
                    return Expr::Const(Value::Error(if self.ctx.xlsx && !external {
                        ErrorKind::Name
                    } else {
                        ErrorKind::Ref
                    }));
                }
            }
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

    /// The arguments of a call of `id` once what the function does with a reference over several
    /// sheets is done (see `functions::threed`), or the expression that the call is instead: the
    /// error the function refuses it with, or what a function that asks about the shape of a
    /// reference answers (`ROWS`: the rows of every sheet, `SHEETS`: the sheets).
    fn over_sheets(
        id: sheetrs_engine::ir::FuncId,
        args: Vec<Expr>,
        over: &[(usize, Vec<Expr>)],
    ) -> Result<Vec<Expr>, Expr> {
        let def = id.def();
        let rects = |areas: &[Expr]| -> Vec<Rect> {
            areas
                .iter()
                .filter_map(|e| match e {
                    Expr::Cell(c) => Some(Rect::single(*c)),
                    Expr::Range(r) => Some(*r),
                    _ => None,
                })
                .collect()
        };
        if let [(0, areas)] = over
            && args.len() == 1
        {
            let rs = rects(areas);
            let number = |n: u32| Err(Expr::Const(Value::Number(f64::from(n))));
            match def.name {
                "ROWS" => return number(rs.iter().map(Rect::rows).sum()),
                "COLUMNS" => return number(rs.iter().map(Rect::cols).sum()),
                "SHEETS" => return number(rs.len() as u32),
                "ROW" => return number(rs.first().map_or(1, |r| r.r0 + 1)),
                "COLUMN" => return number(rs.first().map_or(1, |r| r.c0 + 1)),
                _ => {}
            }
        }
        let mut out = Vec::with_capacity(args.len() + over.len());
        for (i, a) in args.into_iter().enumerate() {
            let Some((_, areas)) = over.iter().find(|(j, _)| *j == i) else {
                out.push(a);
                continue;
            };
            // `IF` and `CHOOSE` hand the reference on to the function that reads their result
            // (`SUM(IF(TRUE,Data:Wide!A1:A3,0))` is 39), so a branch takes the cells of every sheet.
            let handed_on = def.ns == Ns::Core && i >= 1 && matches!(def.name, "IF" | "CHOOSE");
            let outcome = if handed_on {
                functions::Outcome::Iterate
            } else {
                functions::outcome(def.name, def.ns, i)
            };
            match outcome {
                functions::Outcome::Error(kind) => return Err(Expr::Const(Value::Error(kind))),
                functions::Outcome::Iterate if functions::is_variadic(def) && !handed_on => {
                    out.extend(areas.iter().cloned());
                }
                functions::Outcome::Iterate => match functions::lookup(Ns::Core, "_3DA") {
                    Some(array) => out.push(Expr::Call(array, areas.clone())),
                    None => return Err(Expr::Const(Value::Error(ErrorKind::Ref))),
                },
            }
        }
        Ok(out)
    }

    /// The call of a function that has an error written in the formula (`#NAME?`, `#NUM!`…, not one a
    /// calculation gives) among its arguments, when `LibreOffice` answers something other than what the
    /// function does with the arguments one after the other. `None` for any other call. What it answers
    /// (`functions::literal_error` has the tables, `scripts/probe_literal_error.py` the probe):
    ///
    /// 1. the first error a calculation gave in the arguments (`SUM(#NAME?,NA())` is `#N/A`, written
    ///    first or last), for the functions that do not catch it;
    /// 2. else, where the function wants a reference or an array, the first error written where it wants a
    ///    value, and the refusal (`Err:504` for `OFFSET(#NAME?,0,0)`, `Err:502` for the arrays of
    ///    `CORREL`) if there is none;
    /// 3. else what the function does.
    ///
    /// The call is `_LITERR(answer, calculated…)`, whose answer is the first error of `calculated` and
    /// else the answer: the call itself in 3, a constant in 2. The calculated arguments are evaluated
    /// twice, once there and once inside the call: only a call that has an error written in it pays it.
    fn error_where_a_reference_is_wanted(
        &self,
        c: &ast::Call,
        id: sheetrs_engine::ir::FuncId,
        args: &[Expr],
    ) -> Option<Expr> {
        let def = id.def();
        let variadic = functions::is_variadic(def);
        // The series of `FORECAST.ETS…`: a written error there next to a series of several cells is `Err:502`.
        let series_error = |i: usize| {
            functions::ets_series(def.name).is_some_and(|(a, b)| {
                let other = if i == a { b } else { a };
                (i == a || i == b)
                    && (matches!(args.get(other), Some(Expr::Range(r)) if r.rows() * r.cols() > 1)
                        || matches!(args.get(other), Some(Expr::Const(Value::Array(m))) if m.data.len() > 1))
            })
        };
        let wants =
            |i: usize| functions::wants_reference(def.name, def.ns, i, variadic) || series_error(i);
        let refusal = |i: usize| {
            if series_error(i) {
                Some(ErrorKind::Libre(502))
            } else {
                functions::literal_answer(def.name, def.ns, i, variadic, args.len())
            }
        };
        let written: Vec<Option<ErrorKind>> = c
            .args
            .iter()
            .take(args.len())
            .map(|a| match a {
                Some(ast::Expr::Error(text)) => self.written_error(text),
                _ => None,
            })
            .collect();
        let refused = written
            .iter()
            .enumerate()
            .any(|(i, w)| w.is_some() && wants(i));
        // A criterion of an `…IFS` that is a range of several cells: one value for the call, `#VALUE!` where the
        // formula's cell does not meet it, and that before an error written in the call.
        // The arguments are read from the last to the first and the first error stays: the error of a criterion
        // after the last error written is the answer, one before it is not (`SUMIFS(r,r,r,r,#NAME?)` is `#NAME?`).
        let last_written = written.iter().rposition(Option::is_some);
        let criterion_range = |i: usize, a: &Expr| {
            matches!(a, Expr::Range(r) if functions::is_criterion(def.name, i)
                && r.rows() * r.cols() > 1 && Some(i) > last_written)
        };
        let has_criterion_range = args.iter().enumerate().any(|(i, a)| criterion_range(i, a));
        if !refused
            && (written.iter().all(Option::is_none)
                || !(functions::computed_error_first(def.name, def.ns) || has_criterion_range))
        {
            return None;
        }
        // The arguments that can be an error a calculation gave: a calculation, or a cell where a value is
        // wanted.
        let mut calculated = Vec::new();
        for (i, a) in args.iter().enumerate() {
            if written[i].is_some() {
                continue;
            }
            match a {
                Expr::Unary(..) | Expr::Binary(..) | Expr::Call(..) | Expr::Unknown(..) => {
                    calculated.push(a.clone());
                }
                Expr::Cell(_) if !wants(i) => calculated.push(a.clone()),
                // The criterion of an `…IFS` is one value: a range of several cells is met by the cell of the
                // formula, or is `#VALUE!` where it has none, before the error written in the call is looked at.
                _ if criterion_range(i, a) => calculated.push(a.clone()),
                _ => {}
            }
        }
        let answer = if refused {
            // The first error written where a value is wanted, else the refusal of the first reference.
            let value_error = written
                .iter()
                .enumerate()
                .find_map(|(i, w)| w.filter(|_| !wants(i)))
                .or_else(|| {
                    written
                        .iter()
                        .enumerate()
                        .find_map(|(i, w)| w.and(refusal(i)))
                })
                .unwrap_or(ErrorKind::Libre(504));
            Expr::Const(Value::Error(value_error))
        } else if calculated.is_empty() {
            return None;
        } else {
            Expr::Call(id, args.to_vec())
        };
        if calculated.is_empty() {
            return Some(answer);
        }
        let mut out = vec![answer];
        out.extend(calculated);
        Some(Expr::Call(functions::lookup(Ns::Core, "_LITERR")?, out))
    }

    /// The error `text` is when it is one of the six `LibreOffice` reads as an error written in a
    /// formula: in capitals, not a bare `#REF!` of an XLSX (a name there) nor a qualified one
    /// (`Sheet!#REF!`).
    fn written_error(&self, text: &str) -> Option<ErrorKind> {
        if text != text.to_ascii_uppercase() || (self.ctx.xlsx && text == "#REF!") {
            return None;
        }
        match ErrorKind::parse(text)? {
            kind @ (ErrorKind::Null
            | ErrorKind::DivZero
            | ErrorKind::Value
            | ErrorKind::Ref
            | ErrorKind::Name
            | ErrorKind::Num
            | ErrorKind::Na) => Some(kind),
            _ => None,
        }
    }

    fn call(&mut self, c: &ast::Call) -> Expr {
        let name = split_function_name(&c.name);
        if name == "LET" {
            return self.let_call(c);
        }
        if let Some(e) = self.formula_info(c, &name) {
            return e;
        }
        // `LibreOffice` takes an array written in the formula for logical when its first element
        // is `TRUE` or `FALSE` (`ISLOGICAL({TRUE,FALSE})`), and the array a comparison gives for
        // a number.
        if name == "ISLOGICAL"
            && let [Some(ast::Expr::Array(rows))] = c.args.as_slice()
            && matches!(
                rows.first().and_then(|r| r.first()),
                Some(ast::Expr::Bool(_))
            )
        {
            return Expr::Const(Value::Bool(true));
        }
        // The database functions tell `I1` from `I1:I1` in their field and criteria arguments
        // (`LibreOffice`: a cell is no block of criteria, a range of one cell is a block with no
        // condition; a range is never a field), so a range of one cell written as a range stays one.
        let database = DATABASE_FUNCTIONS.contains(&name.as_str());
        // The arguments that are a reference over several sheets, with the area of each sheet.
        let mut over_sheets: Vec<(usize, Vec<Expr>)> = Vec::new();
        let mut args: Vec<Expr> = Vec::with_capacity(c.args.len());
        for (i, a) in c.args.iter().enumerate() {
            args.push(match a {
                None => Expr::Missing,
                Some(a) => {
                    if let Some(areas) = self.three_d(a) {
                        over_sheets.push((i, areas));
                        Expr::Missing
                    } else {
                        match self.expr(a) {
                            Expr::Cell(cell)
                                if database
                                    && matches!(i, 1 | 2)
                                    && matches!(
                                        a,
                                        ast::Expr::Ref(ast::Reference {
                                            area: ast::RefArea::Range(..),
                                            ..
                                        })
                                    ) =>
                            {
                                Expr::Range(Rect::single(cell))
                            }
                            e => e,
                        }
                    }
                }
            });
        }
        if name == "_LONGTEXT"
            && self.ctx.longtext
            && let Some(joined) = join_longtext(&args)
        {
            return Expr::Const(Value::text(joined));
        }
        let (ns, name) = self.function_namespace(&c.name);
        if let Some(id) = functions::lookup(ns, &name) {
            let def = id.def();
            // An error written in the formula where this function reads arrays is an array of one element that
            // is the error: the function does with it what it does with an error element.
            let args: Vec<Expr> = if functions::token_as_array(def.name, def.ns) {
                args.into_iter()
                    .enumerate()
                    .map(|(i, a)| match (&a, c.args.get(i)) {
                        (Expr::Const(Value::Error(kind)), Some(Some(ast::Expr::Error(text))))
                            if self.written_error(text).is_some() =>
                        {
                            Expr::Const(Value::Array(std::sync::Arc::new(
                                sheetrs_engine::value::Array::new(1, 1, vec![Value::Error(*kind)]),
                            )))
                        }
                        _ => a,
                    })
                    .collect()
            } else {
                args
            };
            // `TEXTSPLIT` leaves out an error written in a delimiter or an option (`TEXTSPLIT("a",#NAME?)` is `a`),
            // and answers it when the text is a reference (`TEXTSPLIT(Data!A1,#NAME?)` is `#NAME?`) or an argument
            // before it has several cells (`TEXTSPLIT(1,Data!A1:A3,#NAME?)` is `#NAME?`, `TEXTSPLIT(1,#NAME?,Data!A1:A3)`
            // is 1).
            let args: Vec<Expr> = if def.name == "TEXTSPLIT" {
                let many = |a: &Expr| match a {
                    Expr::Range(r) => r.rows() * r.cols() > 1,
                    Expr::Const(Value::Array(m)) => m.data.len() > 1,
                    _ => false,
                };
                let text_is_reference =
                    matches!(args.first(), Some(Expr::Cell(_) | Expr::Range(_)));
                let before: Vec<bool> = (0..args.len())
                    .map(|i| text_is_reference || args[..i].iter().any(many))
                    .collect();
                args.into_iter()
                    .enumerate()
                    .map(|(i, a)| match (&a, c.args.get(i)) {
                        (Expr::Const(Value::Error(_)), Some(Some(ast::Expr::Error(text))))
                            if i > 0 && !before[i] && self.written_error(text).is_some() =>
                        {
                            Expr::Missing
                        }
                        _ => a,
                    })
                    .collect()
            } else {
                args
            };
            // `FOURIER` leaves out an error written in an option that its data do not look at: the data of one
            // cell take no option (`FOURIER(Data!A1,#NAME?)` is the cell), the data of several cells take the
            // first one (the grouping: `FOURIER(Data!A1:A3,#NAME?)` is `#NAME?`) and no more (`FOURIER(Data!A1:A3,1,#NAME?)`
            // is the transform), and a value written in the formula takes all (`FOURIER(1,#NAME?)` is `#NAME?`).
            let args: Vec<Expr> = if def.name == "FOURIER" {
                // a grouping of several cells makes the others be looked at (`FOURIER(Data!A1:A3,Data!A1:A3,#NAME?)`).
                let grouping_is_many =
                    matches!(args.get(1), Some(Expr::Range(r)) if r.rows() * r.cols() > 1);
                let looked_at = |i: usize| match args.first() {
                    Some(Expr::Range(r)) if r.rows() * r.cols() > 1 => i == 1 || grouping_is_many,
                    Some(Expr::Const(Value::Array(m))) if m.data.len() > 1 => {
                        i == 1 || grouping_is_many
                    }
                    Some(Expr::Cell(_) | Expr::Range(_)) => false,
                    _ => true,
                };
                args.iter()
                    .enumerate()
                    .map(|(i, a)| match (a, c.args.get(i)) {
                        (Expr::Const(Value::Error(_)), Some(Some(ast::Expr::Error(text))))
                            if i > 0 && !looked_at(i) && self.written_error(text).is_some() =>
                        {
                            Expr::Missing
                        }
                        _ => a.clone(),
                    })
                    .collect()
            } else {
                args
            };
            // An argument written empty is what `LibreOffice` makes of it for this function.
            let args: Vec<Expr> = args
                .into_iter()
                .enumerate()
                .map(|(i, a)| match a {
                    // An argument that has to be a range or an array and is written empty is a plain
                    // value (`SUMIF(r, ">3", )` is `Err:504`).
                    Expr::Missing
                        if def
                            .matrix
                            .iter()
                            .any(|&(mask, _)| mask >> i.min(31) & 1 == 1) =>
                    {
                        Expr::Const(Value::Number(0.0))
                    }
                    // In an ODS the last argument of `IFNA` written empty is the number 0 (`IFNA(NA();)` is 0), where an
                    // XLSX drops it (`IFNA(NA(),)` is `#N/A`).
                    Expr::Missing if self.ctx.ods && def.name == "IFNA" && i == 1 => {
                        Expr::Const(Value::Number(0.0))
                    }
                    Expr::Missing => match functions::omitted(def.name, def.ns, i) {
                        functions::Omitted::Absent | functions::Omitted::Dropped => Expr::Missing,
                        functions::Omitted::Zero => Expr::Const(Value::Number(0.0)),
                        functions::Omitted::EmptyText => Expr::Const(Value::text("")),
                        functions::Omitted::Error(kind) => Expr::Const(Value::Error(kind)),
                    },
                    a => a,
                })
                .collect();
            // The last argument written empty is not there for the functions that read it so.
            let mut args = args;
            if over_sheets.is_empty()
                && args.last() == Some(&Expr::Missing)
                && functions::omitted(def.name, def.ns, args.len() - 1)
                    == functions::Omitted::Dropped
            {
                args.pop();
            }
            let args = if over_sheets.is_empty() {
                args
            } else {
                match Self::over_sheets(id, args, &over_sheets) {
                    Ok(args) => args,
                    Err(e) => return e,
                }
            };
            if args.len() < def.min as usize || args.len() > def.max as usize {
                self.diags.push(format!(
                    "{}: wrong number of arguments ({})",
                    id.qualified(),
                    args.len()
                ));
            }
            if over_sheets.is_empty()
                && let Some(e) = self.error_where_a_reference_is_wanted(c, id, &args)
            {
                return e;
            }
            Expr::Call(id, args)
        } else {
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
/// (XLSX: the function is newer than the file format) and the namespaces of `OpenDocument`.
#[must_use]
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
    // The range operator over an array written in the formula is `Err:502`.
    let literal_array = |e: &Expr| matches!(e, Expr::Const(Value::Array(_)));
    if literal_array(&a) || literal_array(&b) {
        return Expr::Const(Value::Error(ErrorKind::Libre(502)));
    }
    // References on two sheets joined by `:` are one over all the sheets from one to the other.
    if let (Some(x), Some(y)) = (rect(&a), rect(&b))
        && x.sheet != y.sheet
    {
        return Expr::Const(Value::Error(ErrorKind::Value));
    }
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

    /// What the call of `f` lowers to, as the error it is when it is one and as the name of the function
    /// and its arguments when it is the answer of `_LITERR`.
    fn literal_call(f: &str, xlsx: bool) -> Expr {
        let (sheets, names) = (vec!["Data".to_string()], HashMap::new());
        let mut c = ctx(&sheets, &names);
        c.xlsx = xlsx;
        parse(f, &c, &mut vec![]).unwrap()
    }

    #[test]
    fn an_error_written_where_a_reference_is_wanted_is_err_504() {
        let libre = |code| Expr::Const(Value::Error(ErrorKind::Libre(code)));
        // `OFFSET(#NAME?,0,0)`, whichever the error and the file.
        for token in ["#NAME?", "#NUM!", "#DIV/0!", "#N/A", "#NULL!", "#VALUE!"] {
            assert_eq!(
                literal_call(&format!("OFFSET({token},0,0)"), true),
                libre(504)
            );
            assert_eq!(
                literal_call(&format!("OFFSET({token},0,0)"), false),
                libre(504)
            );
        }
        // Where the function wants an array the answer is the one of the array (`CORREL`).
        assert_eq!(literal_call("CORREL(#NAME?,Data!A1:A3)", true), libre(502));
        // A bare `#REF!` is a name in an XLSX (`#NAME?` is what OFFSET propagates) and an error in an ODS.
        assert_ne!(literal_call("OFFSET(#REF!,0,0)", true), libre(504));
        assert_eq!(literal_call("OFFSET(#REF!,0,0)", false), libre(504));
        // Where a value is wanted the error is that error, and it wins over the reference that is none.
        assert_eq!(
            literal_call("OFFSET(#NUM!,#NAME?,0)", true),
            Expr::Const(Value::Error(ErrorKind::Name))
        );
        assert_eq!(
            literal_call("OFFSET(Data!A1,#NAME?,0)", true),
            Expr::Call(
                functions::lookup(Ns::Core, "OFFSET").unwrap(),
                vec![
                    Expr::Cell(CellRef::new(0, 0, 0)),
                    Expr::Const(Value::Error(ErrorKind::Name)),
                    n(0.0)
                ]
            )
        );
    }

    #[test]
    fn an_error_a_calculation_gave_wins_over_an_error_written_in_the_formula() {
        // `SUM(#NAME?,NA())` is `#N/A`: the call is `_LITERR(SUM(…), NA())`.
        let Expr::Call(id, args) = literal_call("SUM(#NAME?,NA())", true) else {
            panic!("not a call");
        };
        assert_eq!(id.def().name, "_LITERR");
        assert!(matches!(&args[0], Expr::Call(sum, _) if sum.def().name == "SUM"));
        assert_eq!(args.len(), 2);
        // The calculated error wins over the refusal, as in `OFFSET(#NAME?,NA(),0)`.
        let Expr::Call(id, args) = literal_call("OFFSET(#NAME?,NA(),0)", true) else {
            panic!("not a call");
        };
        assert_eq!(id.def().name, "_LITERR");
        assert_eq!(args[0], Expr::Const(Value::Error(ErrorKind::Libre(504))));
        // A call with an error written and nothing calculated is not rewritten (`SUM(#NAME?)`), and a function
        // that catches the calculated error is not either (`IFERROR(NA(),#NAME?)`).
        assert!(
            matches!(literal_call("SUM(#NAME?,1)", true), Expr::Call(sum, _) if sum.def().name == "SUM")
        );
        assert!(
            matches!(literal_call("IFERROR(1/0,#NAME?)", true), Expr::Call(f, _) if f.def().name == "IFERROR")
        );
    }

    #[test]
    fn a_sheet_that_is_not_there_is_an_unknown_name_in_an_xlsx() {
        let (sheets, names) = (vec!["Data".to_string()], HashMap::new());
        let mut c = ctx(&sheets, &names);
        let name = |c: &ParseCtx, f: &str| match parse(f, c, &mut vec![]).unwrap() {
            Expr::Const(Value::Error(k)) => k,
            other => panic!("{f}: {other:?}"),
        };
        // `LibreOffice` reads `Zzz!A1` of an XLSX as a name that is not defined; an ODS keeps `#REF!`.
        c.xlsx = true;
        assert_eq!(name(&c, "Zzz!A1"), ErrorKind::Name);
        c.xlsx = false;
        assert_eq!(name(&c, "Zzz!A1"), ErrorKind::Ref);
    }

    #[test]
    fn a_reference_over_sheets_is_the_area_of_each_sheet_in_a_call_and_no_value_elsewhere() {
        let sheets: Vec<String> = ["Data", "Mid", "Wide"]
            .iter()
            .map(ToString::to_string)
            .collect();
        let names = HashMap::new();
        let c = ctx(&sheets, &names);
        let areas = |f: &str| match parse(f, &c, &mut vec![]).unwrap() {
            Expr::Call(_, args) => args,
            other => panic!("{f}: {other:?}"),
        };
        // `SUM` has an argument for each sheet; the area is the same on all of them.
        let cell = |s: u16| Expr::Cell(CellRef::new(s, 0, 0));
        assert_eq!(
            areas("SUM(Data!A1:Wide!A1)"),
            vec![cell(0), cell(1), cell(2)]
        );
        assert_eq!(areas("SUM(Data:Mid!A1)"), vec![cell(0), cell(1)]);
        assert_eq!(areas("SUM(Wide:Data!A1)"), vec![cell(0), cell(1), cell(2)]);
        assert_eq!(
            areas("SUM(Data:Mid!A1:B2)"),
            vec![
                Expr::Range(Rect::new(0, 0, 0, 1, 1)),
                Expr::Range(Rect::new(1, 0, 0, 1, 1))
            ]
        );
        // The functions that ask about the shape answer with it.
        assert_eq!(
            parse("ROWS(Data:Wide!A1:A3)", &c, &mut vec![]).unwrap(),
            n(9.0)
        );
        assert_eq!(
            parse("COLUMNS(Data!A1:Wide!B1)", &c, &mut vec![]).unwrap(),
            n(6.0)
        );
        // Anywhere else it is not a value.
        let value = Expr::Const(Value::Error(ErrorKind::Value));
        assert_eq!(parse("Data!A1:Wide!A1", &c, &mut vec![]).unwrap(), value);
        assert_eq!(parse("Data:Wide!A1", &c, &mut vec![]).unwrap(), value);
        // A function that takes one sheet refuses it with its own error.
        let refused = |f: &str| match parse(f, &c, &mut vec![]).unwrap() {
            Expr::Const(Value::Error(k)) => k,
            other => panic!("{f}: {other:?}"),
        };
        assert_eq!(
            refused("SUMIF(Data:Wide!A1:A3,\">3\")"),
            ErrorKind::Libre(504)
        );
        assert_eq!(refused("INDEX(Data:Wide!A1:A3,1)"), ErrorKind::Ref);
    }

    #[test]
    fn a_text_with_an_array_that_is_no_formula_is_the_error_libreoffice_gives() {
        // Probed in LibreOffice 26.8 (`scripts/lo_probe.py`): the error of the cell, not a build error.
        let err = |src: &str| match p(src) {
            Expr::Const(Value::Error(ErrorKind::Libre(code))) => code,
            other => panic!("{src}: {other:?}"),
        };
        for (src, code) in [
            ("{1,2}{3}", 509),
            ("{1,2}}", 508),
            ("{1 2}", 539),
            ("{\"x\",'y'}", 539),
            ("{1,2}{", 512),
            ("{", 512),
            ("{{1}}", 533),
            ("{1,2}:A1", 502),
            ("{1}:{2}", 502),
            ("SUM(1,", 511),
            ("1+", 520),
        ] {
            assert_eq!(err(src), code, "{src}");
        }
        // What is left open is closed, and what is inside is the formula.
        for (open, closed) in [
            ("{1,2", "{1,2}"),
            ("SUM({1,2},{3", "SUM({1,2},{3})"),
            ("SUM(1,2", "SUM(1,2)"),
        ] {
            assert_eq!(p(open), p(closed), "{open}");
        }
    }

    #[test]
    fn unary_minus_and_the_operators_bind_as_in_spreadsheets() {
        // -2^2 = 4: unary minus binds tighter than ^.
        let Expr::Binary(BinOp::Pow, operands) = p("-2^2") else {
            panic!()
        };
        assert_eq!(operands[0], n(-2.0));
        let Expr::Binary(BinOp::Add, operands) = p("1+2*3") else {
            panic!()
        };
        assert!(matches!(operands[1], Expr::Binary(BinOp::Mul, _)));
        assert!(matches!(p("1&2=3"), Expr::Binary(BinOp::Eq, ..)));
        assert!(matches!(p("50%"), Expr::Unary(UnOp::Percent, _)));
    }

    #[test]
    fn cell_range_sheet_whole_row_column_and_name_references_parse() {
        assert_eq!(p("$B$2"), Expr::Cell(CellRef::new(0, 1, 1)));
        assert_eq!(p("'My Sheet'!A1:B2"), Expr::Range(Rect::new(1, 0, 0, 1, 1)));
        assert_eq!(p("Sheet1!C3"), Expr::Cell(CellRef::new(0, 2, 2)));
        assert!(matches!(p("A:A"), Expr::Range(r) if r.rows() == addr::MAX_ROWS));
        assert!(matches!(p("1:2"), Expr::Range(r) if r.cols() == addr::MAX_COLS));
        assert_eq!(p("Rate"), Expr::Cell(CellRef::new(0, 0, 1)));
    }

    #[test]
    fn calls_parse_with_their_namespace_and_the_xlfn_prefix() {
        let Expr::Call(id, args) = p("SUM(A1:A3, 1)") else {
            panic!()
        };
        assert_eq!(id.qualified(), "core::SUM");
        assert_eq!(args.len(), 2);
        let Expr::Call(id, _) = p("_xlfn.CONCAT(\"a\")") else {
            panic!()
        };
        assert_eq!(id.qualified(), "microsoft::CONCAT");
        // An argument written empty is what `LibreOffice` makes of it for the function: the number 0
        // for the then-branch of `IF`, absent (the default) for the digits of `ROUND`.
        assert!(matches!(p("IF(A1,,2)"), Expr::Call(_, a) if a[1] == n(0.0)));
        assert!(matches!(p("ROUND(1.5,)"), Expr::Call(_, a) if a[1] == Expr::Missing));
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
    fn strings_errors_booleans_arrays_and_exponent_numbers_parse_as_constants() {
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
    fn a_range_whose_end_is_a_function_call_is_a_call() {
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

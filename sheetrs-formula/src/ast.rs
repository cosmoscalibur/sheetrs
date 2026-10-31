//! The formula AST. It keeps what a rewriter or a linter needs and a lowering
//! pass would discard: absolute (`$`) markers, the function name as written
//! (prefixes included), the original text of numbers, explicit parentheses and
//! omitted arguments. Whitespace is not kept: printing yields the canonical
//! spelling, and `parse(print(parse(s))) == parse(s)`.

/// A formula body (the text after `=`).
#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    /// A number with its text (`1.50`, `1E3`) so printing does not change it.
    Number(Number),
    Text(String),
    Bool(bool),
    /// An error literal as written (`#N/A`, `#REF!`).
    Error(String),
    /// A constant array `{1,2;3,4}` (rows of constants).
    Array(Vec<Vec<Expr>>),
    Ref(Reference),
    /// A defined name.
    Name(String),
    Call(Call),
    Unary(UnOp, Box<Expr>),
    /// Postfix `%`.
    Percent(Box<Expr>),
    Binary(BinOp, Box<Expr>, Box<Expr>),
    /// `( … )`, kept because it can change meaning and is part of the text.
    Paren(Box<Expr>),
}

#[derive(Debug, Clone, PartialEq)]
pub struct Number {
    pub value: f64,
    pub text: String,
}

impl Number {
    pub fn new(value: f64) -> Self {
        Self {
            text: format_number(value),
            value,
        }
    }
}

fn format_number(v: f64) -> String {
    if v == v.trunc() && v.abs() < 1e15 {
        format!("{}", v as i64)
    } else {
        format!("{v}")
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Call {
    /// As written, e.g. `SUM`, `sum`, `_xlfn.CONCAT`, `COM.MICROSOFT.CONCAT`.
    pub name: String,
    /// `None` is an omitted argument (`IF(A1,,2)`).
    pub args: Vec<Option<Expr>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnOp {
    Neg,
    Plus,
    /// The implicit-intersection operator `@`.
    Implicit,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinOp {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Concat,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
    /// `:` between operands that are not both literal references.
    Range,
}

impl BinOp {
    pub fn symbol(self) -> &'static str {
        match self {
            Self::Add => "+",
            Self::Sub => "-",
            Self::Mul => "*",
            Self::Div => "/",
            Self::Pow => "^",
            Self::Concat => "&",
            Self::Eq => "=",
            Self::Ne => "<>",
            Self::Lt => "<",
            Self::Le => "<=",
            Self::Gt => ">",
            Self::Ge => ">=",
            Self::Range => ":",
        }
    }
}

/// A reference to cells, possibly on another sheet.
#[derive(Debug, Clone, PartialEq)]
pub struct Reference {
    /// Sheet name, unquoted. `None`: the sheet of the formula.
    pub sheet: Option<String>,
    pub area: RefArea,
}

/// A cell address, zero-based, with its `$` markers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellAddr {
    pub row: u32,
    pub col: u32,
    pub row_abs: bool,
    pub col_abs: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ColAddr {
    pub col: u32,
    pub abs: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RowAddr {
    pub row: u32,
    pub abs: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RefArea {
    Cell(CellAddr),
    /// `A1:B2`, in the order written.
    Range(CellAddr, CellAddr),
    /// `A:C`.
    Cols(ColAddr, ColAddr),
    /// `2:5`.
    Rows(RowAddr, RowAddr),
}

impl Expr {
    /// Visit this node and all descendants, parents first.
    pub fn walk(&self, f: &mut impl FnMut(&Expr)) {
        f(self);
        match self {
            Self::Array(rows) => rows.iter().flatten().for_each(|e| e.walk(f)),
            Self::Call(c) => c.args.iter().flatten().for_each(|e| e.walk(f)),
            Self::Unary(_, e) | Self::Percent(e) | Self::Paren(e) => e.walk(f),
            Self::Binary(_, a, b) => {
                a.walk(f);
                b.walk(f);
            }
            _ => {}
        }
    }

    /// Mutable walk, parents first; `f` may replace the node it is given.
    pub fn walk_mut(&mut self, f: &mut impl FnMut(&mut Expr)) {
        f(self);
        match self {
            Self::Array(rows) => rows.iter_mut().flatten().for_each(|e| e.walk_mut(f)),
            Self::Call(c) => c.args.iter_mut().flatten().for_each(|e| e.walk_mut(f)),
            Self::Unary(_, e) | Self::Percent(e) | Self::Paren(e) => e.walk_mut(f),
            Self::Binary(_, a, b) => {
                a.walk_mut(f);
                b.walk_mut(f);
            }
            _ => {}
        }
    }

    /// Every reference in the formula, in source order.
    pub fn references(&self) -> Vec<&Reference> {
        let mut out = Vec::new();
        collect_refs(self, &mut out);
        out
    }
}

fn collect_refs<'a>(e: &'a Expr, out: &mut Vec<&'a Reference>) {
    match e {
        Expr::Ref(r) => out.push(r),
        Expr::Array(rows) => rows.iter().flatten().for_each(|x| collect_refs(x, out)),
        Expr::Call(c) => c.args.iter().flatten().for_each(|x| collect_refs(x, out)),
        Expr::Unary(_, x) | Expr::Percent(x) | Expr::Paren(x) => collect_refs(x, out),
        Expr::Binary(_, a, b) => {
            collect_refs(a, out);
            collect_refs(b, out);
        }
        _ => {}
    }
}

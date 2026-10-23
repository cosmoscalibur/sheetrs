//! Hand-written Pratt parser for the Excel formula syntax as the readers
//! normalize it (ODS references are already converted; `;` is accepted as an
//! argument separator). One pass over the bytes, no token vector.

use crate::addr;
use crate::ast::{BinOp, Call, CellAddr, ColAddr, Expr, Number, RefArea, Reference, RowAddr, UnOp};

#[derive(Debug, Clone, PartialEq)]
pub struct ParseError {
    pub pos: usize,
    pub msg: String,
}

impl std::fmt::Display for ParseError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "at {}: {}", self.pos, self.msg)
    }
}

impl std::error::Error for ParseError {}

const ERRORS: [&str; 10] = [
    "#GETTING_DATA",
    "#DIV/0!",
    "#VALUE!",
    "#SPILL!",
    "#NULL!",
    "#NAME?",
    "#CALC!",
    "#REF!",
    "#NUM!",
    "#N/A",
];

/// Parse a formula; a leading `=` is accepted and dropped.
pub fn parse(src: &str) -> Result<Expr, ParseError> {
    let src = src.strip_prefix('=').unwrap_or(src);
    let mut p = Parser {
        s: src.as_bytes(),
        src,
        i: 0,
    };
    let e = p.expr(0)?;
    p.ws();
    if p.i < p.s.len() {
        return p.fail("unexpected trailing input");
    }
    Ok(e)
}

struct Parser<'a> {
    s: &'a [u8],
    src: &'a str,
    i: usize,
}

fn is_ident_start(b: u8) -> bool {
    b.is_ascii_alphabetic() || matches!(b, b'_' | b'$' | b'\\') || b >= 0x80
}

fn is_ident(b: u8) -> bool {
    b.is_ascii_alphanumeric() || matches!(b, b'_' | b'.' | b'$' | b'\\' | b'?') || b >= 0x80
}

/// Operator, binding power and length.
fn infix(op: &[u8]) -> Option<(BinOp, u8, usize)> {
    Some(match op {
        [b'<', b'=', ..] => (BinOp::Le, 1, 2),
        [b'>', b'=', ..] => (BinOp::Ge, 1, 2),
        [b'<', b'>', ..] => (BinOp::Ne, 1, 2),
        [b'=', ..] => (BinOp::Eq, 1, 1),
        [b'<', ..] => (BinOp::Lt, 1, 1),
        [b'>', ..] => (BinOp::Gt, 1, 1),
        [b'&', ..] => (BinOp::Concat, 2, 1),
        [b'+', ..] => (BinOp::Add, 3, 1),
        [b'-', ..] => (BinOp::Sub, 3, 1),
        [b'*', ..] => (BinOp::Mul, 4, 1),
        [b'/', ..] => (BinOp::Div, 4, 1),
        [b'^', ..] => (BinOp::Pow, 5, 1),
        _ => return None,
    })
}

const PREFIX_BP: u8 = 7;
const RANGE_BP: u8 = 8;

impl Parser<'_> {
    fn fail<T>(&self, msg: &str) -> Result<T, ParseError> {
        Err(ParseError {
            pos: self.i,
            msg: msg.to_string(),
        })
    }

    fn peek(&self) -> Option<u8> {
        self.s.get(self.i).copied()
    }

    fn ws(&mut self) {
        while matches!(self.peek(), Some(b' ' | b'\t' | b'\n' | b'\r')) {
            self.i += 1;
        }
    }

    fn eat(&mut self, b: u8) -> bool {
        self.ws();
        if self.peek() == Some(b) {
            self.i += 1;
            true
        } else {
            false
        }
    }

    fn expr(&mut self, min_bp: u8) -> Result<Expr, ParseError> {
        let mut lhs = self.prefix()?;
        loop {
            self.ws();
            let rest = &self.s[self.i.min(self.s.len())..];
            match rest.first() {
                Some(b'%') => {
                    self.i += 1;
                    lhs = Expr::Percent(Box::new(lhs));
                    continue;
                }
                Some(b':') if RANGE_BP >= min_bp => {
                    self.i += 1;
                    let rhs = self.expr(RANGE_BP + 1)?;
                    lhs = Expr::Binary(BinOp::Range, Box::new(lhs), Box::new(rhs));
                    continue;
                }
                _ => {}
            }
            let Some((op, bp, len)) = infix(rest) else {
                break;
            };
            if bp < min_bp {
                break;
            }
            self.i += len;
            // `^` is left-associative in spreadsheets, like the rest.
            let rhs = self.expr(bp + 1)?;
            lhs = Expr::Binary(op, Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    fn prefix(&mut self) -> Result<Expr, ParseError> {
        self.ws();
        let op = match self.peek() {
            Some(b'-') => UnOp::Neg,
            Some(b'+') => UnOp::Plus,
            Some(b'@') => UnOp::Implicit,
            _ => return self.primary(),
        };
        self.i += 1;
        let e = self.expr(PREFIX_BP)?;
        Ok(Expr::Unary(op, Box::new(e)))
    }

    fn primary(&mut self) -> Result<Expr, ParseError> {
        self.ws();
        let Some(b) = self.peek() else {
            return self.fail("unexpected end of formula");
        };
        match b {
            b'(' => {
                self.i += 1;
                let e = self.expr(0)?;
                if !self.eat(b')') {
                    return self.fail("expected `)`");
                }
                Ok(Expr::Paren(Box::new(e)))
            }
            b'"' => Ok(Expr::Text(self.string()?)),
            b'#' => Ok(Expr::Error(self.error_literal()?)),
            b'{' => self.array_literal(),
            b'\'' => {
                let (name, len) = self.quoted_sheet()?;
                self.i += len;
                self.sheet_ref(name)
            }
            b'0'..=b'9' | b'.' => self.number_or_rows(),
            _ if is_ident_start(b) => self.ident(),
            _ => self.fail("unexpected character"),
        }
    }

    fn string(&mut self) -> Result<String, ParseError> {
        self.i += 1;
        let mut out = String::new();
        let mut start = self.i;
        loop {
            match self.peek() {
                None => return self.fail("unterminated string"),
                Some(b'"') => {
                    out.push_str(&self.src[start..self.i]);
                    self.i += 1;
                    if self.peek() == Some(b'"') {
                        out.push('"');
                        self.i += 1;
                        start = self.i;
                    } else {
                        return Ok(out);
                    }
                }
                Some(_) => self.i += 1,
            }
        }
    }

    fn error_literal(&mut self) -> Result<String, ParseError> {
        let rest = &self.src[self.i..];
        for e in ERRORS {
            if rest.len() >= e.len() && rest[..e.len()].eq_ignore_ascii_case(e) {
                self.i += e.len();
                return Ok(e.to_string());
            }
        }
        self.fail("unknown error literal")
    }

    /// Unquoted name and the byte length consumed through the `!`.
    fn quoted_sheet(&self) -> Result<(String, usize), ParseError> {
        match addr::split_sheet(&self.src[self.i..]) {
            (Some(name), tail) => Ok((name, self.src.len() - self.i - tail.len())),
            _ => self.fail("malformed quoted sheet name"),
        }
    }

    fn number(&mut self) -> Result<Number, ParseError> {
        let start = self.i;
        while matches!(self.peek(), Some(b'0'..=b'9' | b'.')) {
            self.i += 1;
        }
        if matches!(self.peek(), Some(b'e' | b'E')) {
            let save = self.i;
            self.i += 1;
            if matches!(self.peek(), Some(b'+' | b'-')) {
                self.i += 1;
            }
            if matches!(self.peek(), Some(b'0'..=b'9')) {
                while matches!(self.peek(), Some(b'0'..=b'9')) {
                    self.i += 1;
                }
            } else {
                self.i = save;
            }
        }
        let text = &self.src[start..self.i];
        match text.parse() {
            Ok(value) => Ok(Number {
                value,
                text: text.to_string(),
            }),
            Err(_) => self.fail("bad number"),
        }
    }

    fn number_or_rows(&mut self) -> Result<Expr, ParseError> {
        let start = self.i;
        let n = self.number()?;
        // Whole-row reference `1:3`.
        if self.peek() == Some(b':') && matches!(self.s.get(self.i + 1), Some(b'0'..=b'9')) {
            let save = self.i;
            self.i += 1;
            let rhs_start = self.i;
            while matches!(self.peek(), Some(b'0'..=b'9')) {
                self.i += 1;
            }
            if let (Some((ra, _)), Some((rb, _))) = (
                addr::parse_row(&self.src[start..save]),
                addr::parse_row(&self.src[rhs_start..self.i]),
            ) {
                return Ok(Expr::Ref(Reference {
                    sheet: None,
                    area: RefArea::Rows(
                        RowAddr {
                            row: ra,
                            abs: false,
                        },
                        RowAddr {
                            row: rb,
                            abs: false,
                        },
                    ),
                }));
            }
            self.i = save;
        }
        Ok(Expr::Number(n))
    }

    fn array_literal(&mut self) -> Result<Expr, ParseError> {
        self.i += 1; // `{`
        let mut rows: Vec<Vec<Expr>> = vec![vec![]];
        loop {
            self.ws();
            let v = match self.peek() {
                Some(b'"') => Expr::Text(self.string()?),
                Some(b'#') => Expr::Error(self.error_literal()?),
                Some(b'-' | b'+' | b'0'..=b'9' | b'.') => {
                    let sign = match self.peek() {
                        Some(b'-') => Some(UnOp::Neg),
                        Some(b'+') => Some(UnOp::Plus),
                        _ => None,
                    };
                    if sign.is_some() {
                        self.i += 1;
                    }
                    let n = Expr::Number(self.number()?);
                    match sign {
                        Some(op) => Expr::Unary(op, Box::new(n)),
                        None => n,
                    }
                }
                Some(c) if is_ident_start(c) => {
                    let start = self.i;
                    while self.peek().is_some_and(is_ident) {
                        self.i += 1;
                    }
                    match self.src[start..self.i].to_ascii_uppercase().as_str() {
                        "TRUE" => Expr::Bool(true),
                        "FALSE" => Expr::Bool(false),
                        _ => return self.fail("array literals hold constants only"),
                    }
                }
                _ => return self.fail("bad array literal"),
            };
            if let Some(r) = rows.last_mut() {
                r.push(v);
            }
            self.ws();
            match self.peek() {
                Some(b',') => self.i += 1,
                Some(b';') => {
                    self.i += 1;
                    rows.push(vec![]);
                }
                Some(b'}') => {
                    self.i += 1;
                    break;
                }
                _ => return self.fail("expected `,`, `;` or `}`"),
            }
        }
        let cols = rows[0].len();
        if rows.iter().any(|r| r.len() != cols) {
            return self.fail("ragged array literal");
        }
        Ok(Expr::Array(rows))
    }

    fn ident(&mut self) -> Result<Expr, ParseError> {
        let start = self.i;
        while self.peek().is_some_and(is_ident) {
            self.i += 1;
        }
        let text = &self.src[start..self.i];
        match self.peek() {
            Some(b'(') => {
                self.i += 1;
                return self.call(text);
            }
            Some(b'!') => {
                self.i += 1;
                return self.sheet_ref(text.to_string());
            }
            _ => {}
        }
        match text.to_ascii_uppercase().as_str() {
            "TRUE" => return Ok(Expr::Bool(true)),
            "FALSE" => return Ok(Expr::Bool(false)),
            _ => {}
        }
        if let Some(area) = self.area_from(text) {
            return Ok(Expr::Ref(Reference { sheet: None, area }));
        }
        Ok(Expr::Name(text.to_string()))
    }

    /// Try `text` (and a following `:ref` tail) as an area.
    fn area_from(&mut self, text: &str) -> Option<RefArea> {
        if self.peek() == Some(b':') {
            // Only consume the tail when it forms a literal area; otherwise
            // `:` stays a range operator (`A1:INDEX(…)`).
            let save = self.i;
            self.i += 1;
            let tail_start = self.i;
            while self.peek().is_some_and(is_ident) {
                self.i += 1;
            }
            let mut tail = &self.src[tail_start..self.i];
            if self.peek() == Some(b'!') {
                // `A1:Sheet1!B2`: the second qualifier is redundant.
                self.i += 1;
                let t2 = self.i;
                while self.peek().is_some_and(is_ident) {
                    self.i += 1;
                }
                tail = &self.src[t2..self.i];
            }
            if self.peek() != Some(b'(')
                && let Some(area) = pair_area(text, tail)
            {
                return Some(area);
            }
            self.i = save;
        }
        let (row, col, row_abs, col_abs) = addr::parse_cell(text)?;
        Some(RefArea::Cell(CellAddr {
            row,
            col,
            row_abs,
            col_abs,
        }))
    }

    fn sheet_ref(&mut self, sheet: String) -> Result<Expr, ParseError> {
        let start = self.i;
        while self.peek().is_some_and(is_ident) {
            self.i += 1;
        }
        let text = &self.src[start..self.i];
        match self.area_from(text) {
            Some(area) => Ok(Expr::Ref(Reference {
                sheet: Some(sheet),
                area,
            })),
            // `Sheet1!Name`: a sheet-scoped defined name, kept qualified.
            None => Ok(Expr::Name(format!(
                "{}!{text}",
                crate::print::sheet_text(&sheet)
            ))),
        }
    }

    fn call(&mut self, name: &str) -> Result<Expr, ParseError> {
        let mut args = Vec::new();
        self.ws();
        if self.peek() == Some(b')') {
            self.i += 1;
        } else {
            loop {
                self.ws();
                if matches!(self.peek(), Some(b',' | b';' | b')')) {
                    args.push(None);
                } else {
                    args.push(Some(self.expr(0)?));
                }
                self.ws();
                match self.peek() {
                    Some(b',' | b';') => self.i += 1,
                    Some(b')') => {
                        self.i += 1;
                        break;
                    }
                    _ => return self.fail("expected `,` or `)` in call"),
                }
            }
        }
        Ok(Expr::Call(Call {
            name: name.to_string(),
            args,
        }))
    }
}

/// `A1:B2`, `A:C` or `2:5` from the two ends as written.
fn pair_area(a: &str, b: &str) -> Option<RefArea> {
    if let (Some((ra, ca, rea, cea)), Some((rb, cb, reb, ceb))) =
        (addr::parse_cell(a), addr::parse_cell(b))
    {
        return Some(RefArea::Range(
            CellAddr {
                row: ra,
                col: ca,
                row_abs: rea,
                col_abs: cea,
            },
            CellAddr {
                row: rb,
                col: cb,
                row_abs: reb,
                col_abs: ceb,
            },
        ));
    }
    if let (Some((ca, aa)), Some((cb, ab))) = (addr::parse_col(a), addr::parse_col(b)) {
        return Some(RefArea::Cols(
            ColAddr { col: ca, abs: aa },
            ColAddr { col: cb, abs: ab },
        ));
    }
    if let (Some((ra, aa)), Some((rb, ab))) = (addr::parse_row(a), addr::parse_row(b)) {
        return Some(RefArea::Rows(
            RowAddr { row: ra, abs: aa },
            RowAddr { row: rb, abs: ab },
        ));
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    fn p(s: &str) -> Expr {
        parse(s).unwrap()
    }

    fn rt(s: &str) -> String {
        p(s).to_string()
    }

    #[test]
    fn keeps_what_a_rewriter_needs() {
        // Absolute markers, function name as written, number text, parentheses,
        // omitted arguments.
        assert_eq!(rt("$A$1+B$2"), "$A$1+B$2");
        assert_eq!(rt("_xlfn.CONCAT(\"a\")"), "_xlfn.CONCAT(\"a\")");
        assert_eq!(rt("sum(1.50,1E3)"), "sum(1.50,1E3)");
        assert_eq!(rt("(1+2)*3"), "(1+2)*3");
        assert_eq!(rt("IF(A1,,2)"), "IF(A1,,2)");
        assert_eq!(rt("'My Sheet'!A1:B2"), "'My Sheet'!A1:B2");
        assert_eq!(rt("Sheet1!$A:$B+2:3"), "Sheet1!$A:$B+2:3");
        assert_eq!(rt("\"a\"\"b\"&#N/A"), "\"a\"\"b\"&#N/A");
        assert_eq!(rt("{1,2;-3,\"x\"}"), "{1,2;-3,\"x\"}");
        assert_eq!(rt("50%+@A1"), "50%+@A1");
        assert_eq!(rt("A1:INDEX(B1:B5,2)"), "A1:INDEX(B1:B5,2)");
    }

    #[test]
    fn precedence_matches_spreadsheets() {
        // Unary minus binds tighter than `^`.
        let Expr::Binary(BinOp::Pow, l, _) = p("-2^2") else {
            panic!()
        };
        assert!(matches!(*l, Expr::Unary(UnOp::Neg, _)));
        let Expr::Binary(BinOp::Add, _, r) = p("1+2*3") else {
            panic!()
        };
        assert!(matches!(*r, Expr::Binary(BinOp::Mul, ..)));
        assert!(matches!(p("1&2=3"), Expr::Binary(BinOp::Eq, ..)));
        assert!(matches!(p("50%"), Expr::Percent(_)));
        // `^` is left-associative.
        let Expr::Binary(BinOp::Pow, l, _) = p("2^3^2") else {
            panic!()
        };
        assert!(matches!(*l, Expr::Binary(BinOp::Pow, ..)));
    }

    #[test]
    fn references_names_and_booleans() {
        assert!(matches!(p("Rate"), Expr::Name(n) if n == "Rate"));
        assert!(matches!(p("true"), Expr::Bool(true)));
        assert!(matches!(p("Sheet1!A1:Sheet1!B2"), Expr::Ref(_)));
        let Expr::Ref(r) = p("'It''s'!$B$2") else {
            panic!()
        };
        assert_eq!(r.sheet.as_deref(), Some("It's"));
        assert_eq!(r.to_string(), "'It''s'!$B$2");
        assert_eq!(p("A1;B1".trim_end_matches(";B1")).references().len(), 1);
    }

    #[test]
    fn print_parse_is_a_fixed_point() {
        for f in [
            "SUM(A1:A3, 1)",
            "IF( A1 > 0 , \"x\" , \"y\" )",
            "-(-3)",
            "VLOOKUP(K10=\"a\";B3:C6;2)",
            "Data!A1+'Q1 2024'!B2*C:C",
            "$A$1:$B$2",
        ] {
            let once = p(f);
            assert_eq!(p(&once.to_string()), once, "{f}");
        }
    }

    #[test]
    fn errors_have_positions() {
        for bad in [
            "1+",
            "(1",
            "SUM(1,",
            "\"open",
            "[1]Sheet1!B3",
            "A1 B1",
            "{1,2;3}",
        ] {
            assert!(parse(bad).is_err(), "{bad}");
        }
        assert_eq!(parse("1+").unwrap_err().pos, 2);
    }
}

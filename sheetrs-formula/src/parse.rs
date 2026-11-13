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

/// A formula together with its source text. The AST drops whitespace and
/// spelling details; the source and the span of every reference let a consumer
/// get them back: a rule can look at the text between tokens, and an edit can
/// replace just the references that change ([`Formula::rewrite`]).
#[derive(Debug, Clone, PartialEq)]
pub struct Formula {
    /// The text as given, including a leading `=` if there was one.
    pub source: String,
    pub expr: Expr,
    /// Byte range in `source` of each reference, in source order (the order of
    /// [`Expr::references`]).
    pub ref_spans: Vec<std::ops::Range<usize>>,
}

/// Parse a formula; a leading `=` is accepted and dropped.
///
/// # Errors
///
/// When `src` is not a formula.
pub fn parse(src: &str) -> Result<Expr, ParseError> {
    parse_spanned(src).map(|f| f.expr)
}

/// Parse keeping the source and the span of each reference.
///
/// # Errors
///
/// When `source` is not a formula.
pub fn parse_spanned(source: &str) -> Result<Formula, ParseError> {
    let offset = usize::from(source.starts_with('='));
    let body = &source[offset..];
    let mut p = Parser {
        s: body.as_bytes(),
        src: body,
        i: 0,
        spans: Vec::new(),
    };
    let expr = p.expr(0)?;
    p.ws();
    if p.i < p.s.len() {
        return p.fail("unexpected trailing input");
    }
    let ref_spans = p
        .spans
        .iter()
        .map(|&(a, b)| a + offset..b + offset)
        .collect();
    Ok(Formula {
        source: source.to_string(),
        expr,
        ref_spans,
    })
}

struct Parser<'a> {
    s: &'a [u8],
    src: &'a str,
    i: usize,
    /// Spans of the references parsed so far, in source order.
    spans: Vec<(usize, usize)>,
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
        let start = self.i;
        let e = self.primary_at()?;
        if matches!(e, Expr::Ref(_)) {
            self.spans.push((start, self.i));
        }
        Ok(e)
    }

    fn primary_at(&mut self) -> Result<Expr, ParseError> {
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
                return Ok(rest[..e.len()].to_string());
            }
        }
        self.fail("unknown error literal")
    }

    /// Unquoted name and the byte length consumed through the `!`.
    fn quoted_sheet(&self) -> Result<(String, usize), ParseError> {
        let rest = &self.src[self.i..];
        if let (Some(name), tail) = addr::split_sheet(rest) {
            return Ok((name, self.src.len() - self.i - tail.len()));
        }
        // `'Data':'Wide'!A1` and `'Data':Wide!A1` (LibreOffice reads them as `Data:Wide!A1`): the first
        // name is quoted and alone, so what follows it names the last sheet.
        let bad = || self.fail("malformed quoted sheet name");
        let Some(close) = quoted_end(rest) else {
            return bad();
        };
        let first = rest[1..close - 1].replace("''", "'");
        let Some(after) = rest[close..].strip_prefix(':') else {
            return bad();
        };
        let (second, tail) = if after.starts_with('\'') {
            match addr::split_sheet(after) {
                (Some(name), tail) => (name, tail),
                _ => return bad(),
            }
        } else {
            // An unquoted name runs to the next `!`.
            let end = after
                .find(|c: char| !(c.is_alphanumeric() || c == '_' || c == '.'))
                .unwrap_or(after.len());
            match after[end..].strip_prefix('!') {
                Some(tail) if end > 0 => (after[..end].to_string(), tail),
                _ => return bad(),
            }
        };
        Ok((format!("{first}:{second}"), rest.len() - tail.len()))
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

    /// `{1,2;3,4}`. An element is any expression: a spreadsheet only takes constants there, and
    /// what is not one (`{1;A1}`, `{1,2+3}`) is an error when the formula is compiled, which is not
    /// a matter of syntax. The rows are not checked to be of one width either. An element left
    /// out (`{}`, `{1,2,}`, `{1,;2,3}`) is an empty element, written `Expr::Error("")`.
    fn array_literal(&mut self) -> Result<Expr, ParseError> {
        self.i += 1; // `{`
        let mut rows: Vec<Vec<Expr>> = vec![vec![]];
        loop {
            self.ws();
            let v = match self.peek() {
                None => return self.fail("bad array literal"),
                Some(b',' | b';' | b'}') => Expr::Error(String::new()),
                Some(_) => self.expr(0)?,
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
            // `T:V!A1`, the reference from sheet `T` to sheet `V` (the same area on each sheet in
            // between): the sheet is written `T:V`, as the file keeps it when the names need quotes.
            Some(b':') if addr::parse_cell(text).is_none() => {
                // `T:'V'!A1`: the last sheet quoted.
                if self.s.get(self.i + 1) == Some(&b'\'')
                    && let (Some(last), tail) = addr::split_sheet(&self.src[self.i + 1..])
                {
                    let sheet = format!("{text}:{last}");
                    self.i = self.src.len() - tail.len();
                    return self.sheet_ref(sheet);
                }
                let mut j = self.i + 1;
                while self.s.get(j).copied().is_some_and(is_ident) {
                    j += 1;
                }
                if j > self.i + 1 && self.s.get(j) == Some(&b'!') {
                    let sheet = format!("{text}:{}", &self.src[self.i + 1..j]);
                    self.i = j + 1;
                    return self.sheet_ref(sheet);
                }
            }
            _ => {}
        }
        match text.to_ascii_uppercase().as_str() {
            "TRUE" => return Ok(Expr::Bool(true)),
            "FALSE" => return Ok(Expr::Bool(false)),
            _ => {}
        }
        if let Some(area) = self.area_from(text, None) {
            return Ok(Expr::Ref(Reference { sheet: None, area }));
        }
        Ok(Expr::Name(text.to_string()))
    }

    /// Try `text` (and a following `:ref` tail) as an area.
    fn area_from(&mut self, text: &str, sheet: Option<&str>) -> Option<RefArea> {
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
            // `T:V!A:A` is the 3D reference from sheet `T` to sheet `V`, whatever the names look
            // like; `A1:Sheet1!B2` has a redundant second qualifier, and only a cell is followed
            // by it.
            let three_d = self.peek() == Some(b'!') && addr::parse_cell(text).is_none();
            if !three_d && self.peek() == Some(b'!') {
                // `Data!A1:Wide!A1` goes from a sheet to another: not a range of the first one (the
                // `:` stays an operator between two references, which the lowering reads as one
                // over several sheets).
                if sheet.is_some_and(|s| !addr::same_name(s, tail)) {
                    self.i = save;
                    let (row, col, row_abs, col_abs) = addr::parse_cell(text)?;
                    return Some(RefArea::Cell(CellAddr {
                        row,
                        col,
                        row_abs,
                        col_abs,
                    }));
                }
                self.i += 1;
                let t2 = self.i;
                while self.peek().is_some_and(is_ident) {
                    self.i += 1;
                }
                tail = &self.src[t2..self.i];
            }
            if !three_d
                && self.peek() != Some(b'(')
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
        // `Sheet1!#REF!`: what a deleted cell leaves, qualified.
        if self.peek() == Some(b'#') {
            let error = self.error_literal()?;
            return Ok(Expr::Error(format!(
                "{}!{error}",
                crate::print::sheet_text(&sheet)
            )));
        }
        let start = self.i;
        while self.peek().is_some_and(is_ident) {
            self.i += 1;
        }
        let text = &self.src[start..self.i];
        match self.area_from(text, Some(&sheet)) {
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

/// The byte length of the quoted sheet name at the start of `text` (`'Data'` is 6), with `''` for a quote.
fn quoted_end(text: &str) -> Option<usize> {
    let b = text.as_bytes();
    if b.first() != Some(&b'\'') {
        return None;
    }
    let mut i = 1;
    while i < b.len() {
        if b[i] == b'\'' {
            if b.get(i + 1) == Some(&b'\'') {
                i += 2;
                continue;
            }
            return Some(i + 1);
        }
        i += 1;
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
    fn a_reference_from_a_sheet_to_another() {
        // `T:V!A1` is one reference on the sheets from `T` to `V`; written with the sheet in each
        // end it is two references joined by the range operator, not a range of the first sheet.
        let Expr::Ref(r) = p("Data:Wide!A1:B2") else {
            panic!()
        };
        assert_eq!(r.sheet.as_deref(), Some("Data:Wide"));
        let Expr::Ref(r) = p("'My Sheet:Other'!A1") else {
            panic!()
        };
        assert_eq!(r.sheet.as_deref(), Some("My Sheet:Other"));
        let Expr::Binary(BinOp::Range, a, b) = p("Data!A1:Wide!B2") else {
            panic!()
        };
        assert!(matches!(*a, Expr::Ref(Reference { sheet: Some(ref s), .. }) if s == "Data"));
        assert!(matches!(*b, Expr::Ref(Reference { sheet: Some(ref s), .. }) if s == "Wide"));
        // The same sheet on both ends is the range of that sheet.
        let Expr::Ref(r) = p("Data!A1:Data!B2") else {
            panic!()
        };
        assert!(matches!(r.area, RefArea::Range(..)) && r.sheet.as_deref() == Some("Data"));
    }

    #[test]
    fn a_sheet_quoted_at_one_end_or_both_of_a_reference_over_sheets() {
        for text in [
            "'Data':'Wide'!A1",
            "'Data':Wide!A1",
            "Data:'Wide'!A1",
            "'Data':'My ''Sheet'!A1",
        ] {
            let Expr::Ref(r) = p(text) else {
                panic!("{text}")
            };
            let sheet = r.sheet.unwrap();
            assert!(sheet.starts_with("Data:"), "{text}: {sheet}");
        }
        let Expr::Ref(r) = p("'Data':'My ''Sheet'!A1") else {
            panic!()
        };
        assert_eq!(r.sheet.as_deref(), Some("Data:My 'Sheet"));
        assert!(parse("'Data':!A1").is_err());
        assert!(parse("'Data'").is_err());
    }

    #[test]
    fn a_deleted_cell_is_a_qualified_ref_error() {
        assert_eq!(rt("BIENES!#REF!+1"), "BIENES!#REF!+1");
        assert_eq!(rt("SUM('My Sheet'!#REF!,A1)"), "SUM('My Sheet'!#REF!,A1)");
        assert_eq!(rt("#REF!"), "#REF!");
        // It holds no reference to rewrite.
        let f = parse_spanned("BIENES!#REF!+A1").unwrap();
        assert_eq!(f.ref_spans.len(), 1);
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
    fn a_name_before_a_sheet_qualifier_is_a_3d_reference_not_a_column_range() {
        // `T` and `V` are columns, and also the sheets of the span.
        let e = parse("SUM(T:V!A:A)").unwrap();
        let Expr::Call(call) = &e else {
            panic!("{e:?}")
        };
        let Some(Expr::Ref(r)) = &call.args[0] else {
            panic!("{e:?}")
        };
        assert_eq!(r.sheet.as_deref(), Some("T:V"));
        assert_eq!(e.to_string(), "SUM(T:V!A:A)");
        // The qualifier of the second end of a range of cells is redundant.
        let e = parse("A1:Sheet1!B2").unwrap();
        assert!(
            matches!(
                e,
                Expr::Ref(Reference {
                    sheet: None,
                    area: RefArea::Range(..)
                })
            ),
            "{e:?}"
        );
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
            "{1,",
            "{1 2}",
            "{",
        ] {
            assert!(parse(bad).is_err(), "{bad}");
        }
        assert_eq!(parse("1+").unwrap_err().pos, 2);
    }

    #[test]
    fn an_array_takes_any_expression_and_any_width() {
        // What a spreadsheet refuses there is an error of the formula, not of its syntax.
        for ok in [
            "{1;A1}",
            "{1,2+3}",
            "{1,SUM(1,2)}",
            "{1,2;3}",
            "{1,-2,+3}",
            "{1,\"a\"&\"b\"}",
        ] {
            assert!(matches!(p(ok), Expr::Array(_)), "{ok}");
        }
        let Expr::Array(rows) = p("{1,2;3}") else {
            panic!("an array")
        };
        assert_eq!(rows.iter().map(Vec::len).collect::<Vec<_>>(), [2, 1]);
        assert_eq!(rt("{1,A1;2+3,4}"), "{1,A1;2+3,4}");
    }

    #[test]
    fn an_array_may_leave_elements_out() {
        let widths = |src: &str| {
            let Expr::Array(rows) = p(src) else {
                panic!("an array")
            };
            rows.iter().map(Vec::len).collect::<Vec<_>>()
        };
        assert_eq!(widths("{}"), [1]);
        assert_eq!(widths("{1,2,}"), [3]);
        assert_eq!(widths("{1,;2,3}"), [2, 2]);
        assert_eq!(widths("{,1}"), [2]);
        assert_eq!(widths("{1,2;}"), [2, 1]);
        assert_eq!(rt("{1,;2,3}"), "{1,;2,3}");
        assert_eq!(rt("{}"), "{}");
        // A nested array is an element that is not a constant.
        assert!(matches!(p("{{1,2}}"), Expr::Array(_)));
    }
}

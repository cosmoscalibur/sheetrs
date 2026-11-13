//! Reading a formula written in `LibreOffice`'s notation, the way a person types it into a cell
//! (`=SUM($'My Sheet'.A1:B2,C3)`), and writing it in the notation a file keeps: Excel's (`XLSX`)
//! or `OpenFormula`'s (`ODS`, `[$'My Sheet'.A1:.B2]` with `;` between arguments).
//!
//! It is the inverse of [`calc_print`](crate::calc_print) for the `en-US` notation (`,` between
//! arguments and between the columns of an array, `;` between its rows; `;` between arguments is
//! accepted as well, as `LibreOffice` accepts it). Only the surface is translated: references,
//! separators and array constants. Text, numbers, operators, names and functions are copied as they
//! were written, and the result is checked with [`parse`](crate::parse()), so a text that is not a
//! formula is an error here instead of a file `LibreOffice` shows as `Err:5xx`.

use crate::print::sheet_text;

/// The notation to write.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Notation {
    /// `Sheet!A1:B2`, `,`.
    Excel,
    /// `[Sheet.A1:.B2]`, `;`, arrays with `;` between columns and `|` between rows.
    OpenFormula,
}

/// What is wrong with a formula that was typed.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FromCalcError {
    /// The text is not a formula.
    NotAFormula(String),
    /// A reference over several sheets (`$A.A1:$C.A1`), which is not supported.
    ThreeD(String),
}

impl std::fmt::Display for FromCalcError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NotAFormula(why) => write!(f, "the text is not a formula: {why}"),
            Self::ThreeD(r) => write!(
                f,
                "the reference `{r}` spans several sheets, which is not supported"
            ),
        }
    }
}

impl std::error::Error for FromCalcError {}

type Sheet = Option<(bool, String)>;

/// A reference read: its sheet, its first part, its second part and where it ends.
type Found = (Sheet, String, Option<String>, usize);

struct Reader<'a> {
    s: &'a [char],
}

fn is_ident(c: char) -> bool {
    c.is_alphanumeric() || c == '_'
}

impl Reader<'_> {
    fn at(&self, i: usize) -> Option<char> {
        self.s.get(i).copied()
    }

    /// `$'My Sheet'.` or `$Sheet.` or `Sheet.` at `i`: the sheet and the position after the dot.
    fn sheet(&self, mut i: usize) -> Option<(Sheet, usize)> {
        let absolute = self.at(i) == Some('$');
        if absolute {
            i += 1;
        }
        let name = if self.at(i) == Some('\'') {
            i += 1;
            let mut name = String::new();
            loop {
                match self.at(i)? {
                    '\'' if self.at(i + 1) == Some('\'') => {
                        name.push('\'');
                        i += 2;
                    }
                    '\'' => {
                        i += 1;
                        break;
                    }
                    c => {
                        name.push(c);
                        i += 1;
                    }
                }
            }
            name
        } else {
            let start = i;
            while self.at(i).is_some_and(is_ident) {
                i += 1;
            }
            if i == start {
                return None;
            }
            self.s[start..i].iter().collect()
        };
        if self.at(i) != Some('.') {
            return None;
        }
        Some((Some((absolute, name)), i + 1))
    }

    /// A cell, a column or a row of a reference at `i`: `$A$1`, `A`, `3`.
    fn part(&self, mut i: usize) -> Option<(String, usize, Kind)> {
        let start = i;
        if self.at(i) == Some('$') {
            i += 1;
        }
        let letters = i;
        while self.at(i).is_some_and(|c| c.is_ascii_alphabetic()) {
            i += 1;
        }
        let n_letters = i - letters;
        if n_letters > 3 {
            return None;
        }
        let kind = if n_letters > 0 {
            if self.at(i) == Some('$') {
                i += 1;
            }
            let digits = i;
            while self.at(i).is_some_and(|c| c.is_ascii_digit()) {
                i += 1;
            }
            if i > digits {
                Kind::Cell
            } else if self.s[start..i].contains(&'$') && self.at(i - 1) == Some('$') {
                return None;
            } else {
                Kind::Column
            }
        } else {
            let digits = i;
            while self.at(i).is_some_and(|c| c.is_ascii_digit()) {
                i += 1;
            }
            if i == digits {
                return None;
            }
            Kind::Row
        };
        Some((self.s[start..i].iter().collect(), i, kind))
    }

    /// A reference at `i`: `(sheet, first, second, end)`.
    fn reference(&self, i: usize) -> Result<Option<Found>, FromCalcError> {
        if i > 0 && (is_ident(self.s[i - 1]) || matches!(self.s[i - 1], '.' | '$' | '\'')) {
            return Ok(None);
        }
        let (sheet, j) = self.sheet(i).unwrap_or((None, i));
        // A sheet prefix without a cell after it is not one.
        let (sheet, j) = if sheet.is_some() && self.part(j).is_none() {
            (None, i)
        } else {
            (sheet, j)
        };
        let Some((first, mut end, kind)) = self.part(j) else {
            return Ok(None);
        };
        let mut second = None;
        if self.at(end) == Some(':') {
            let after = end + 1;
            let (sheet2, k) = self.sheet(after).unwrap_or((None, after));
            if let Some((b, e, kind2)) = self.part(k)
                && kind2 == kind
            {
                if let Some((_, n2)) = &sheet2
                    && sheet.as_ref().is_none_or(|(_, n1)| !same(n1, n2))
                {
                    let text: String = self.s[i..e].iter().collect();
                    return Err(FromCalcError::ThreeD(text));
                }
                second = Some(b);
                end = e;
            }
        }
        if second.is_none() && kind != Kind::Cell {
            return Ok(None);
        }
        // Whole word: not a name or a function that happens to start like a cell.
        if self
            .at(end)
            .is_some_and(|c| is_ident(c) || c == '(' || c == '.')
        {
            return Ok(None);
        }
        Ok(Some((sheet, first, second, end)))
    }
}

fn same(a: &str, b: &str) -> bool {
    a.to_lowercase() == b.to_lowercase()
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Cell,
    Column,
    Row,
}

/// A formula as a file keeps it (`native`, written in `from`) in the notation `LibreOffice` shows
/// for it (`en-US`: `,` between arguments), without the leading `=`.
///
/// What the application shows depends on how the file wrote the formula, not only on what it means:
///
/// - An XLSX has only sheets that are absolute: `T!A1` is `$T.A1`, also in a formula of the sheet `T`
///   itself, which keeps the sheet it was written with.
/// - An ODS keeps `[.A1]` (a reference without a sheet, shown `A1`), `[T.A1]` (a sheet that is
///   relative, shown `T.A1`), `[$T.A1]` (absolute, shown `$T.A1`) and `[T.A1:T.B3]` (a range that
///   names its sheet twice, shown so): the text is the one of the file with the brackets and the
///   empty sheet taken off.
/// - A function the application does not have is shown as it was written (`foo()`); `known` says
///   whether it has a function of an uppercase name.
#[must_use]
pub fn native_to_calc(native: &str, from: Notation, known: &dyn Fn(&str) -> bool) -> String {
    match from {
        Notation::Excel => excel_to_calc(native, known).unwrap_or_else(|| native.to_string()),
        Notation::OpenFormula => openformula_to_calc(native, known),
    }
}

/// Whether `formula` (written in the `notation` of a file) has the error `#REF!` as an operand of
/// its own, a "bare" one (`#REF!+5`, `SUM(#REF!,1)`): Excel writes it so when the cell a formula
/// read is deleted. `LibreOffice` reads it from an XLSX as the name `#ref!`, which it does not know
/// (`#NAME?`); from an ODS it reads `#ref!` as the error itself (`#REF!`). What is part of a
/// reference (`Sheet!#REF!`, `#REF!!A1`, `[#REF!]`) is not bare, and neither are text, the quoted
/// name of a sheet or, in an ODS, `#ref!` glued to a name (`#ref!a1`, which is an unknown name).
#[must_use]
pub fn has_bare_ref_error(formula: &str, notation: Notation) -> bool {
    if !formula.contains('#') {
        return false;
    }
    let b = formula.as_bytes();
    let ods = notation == Notation::OpenFormula;
    let mut i = 0;
    while i < b.len() {
        match b[i] {
            quote @ (b'"' | b'\'') => {
                i += 1;
                while i < b.len() {
                    if b[i] == quote {
                        if b.get(i + 1) == Some(&quote) {
                            i += 2;
                            continue;
                        }
                        i += 1;
                        break;
                    }
                    i += 1;
                }
            }
            b'[' if ods => {
                // A reference of `OpenFormula`, which may hold a quoted file name.
                i += 1;
                let mut quoted = false;
                while i < b.len() {
                    match b[i] {
                        b'\'' => quoted = !quoted,
                        b']' if !quoted => {
                            i += 1;
                            break;
                        }
                        _ => {}
                    }
                    i += 1;
                }
            }
            b'#' if b.len() >= i + 5 && b[i..i + 5].eq_ignore_ascii_case(b"#REF!") => {
                let qualified = (i > 0 && b[i - 1] == b'!') || b.get(i + 5) == Some(&b'!');
                let glued = ods
                    && b.get(i + 5)
                        .is_some_and(|c| c.is_ascii_alphanumeric() || *c == b'_' || *c == b'.');
                if !qualified && !glued {
                    return true;
                }
                i += 5;
            }
            _ => i += 1,
        }
    }
    false
}

/// An Excel text in the notation `LibreOffice` shows, keeping the spacing it was written with:
/// the references, the names of the functions, the numbers and the constants `TRUE` and `FALSE` are
/// rewritten where they stand, and the rest is copied. `None` when the text is not a formula.
fn excel_to_calc(native: &str, known: &dyn Fn(&str) -> bool) -> Option<String> {
    let formula = crate::parse::parse_spanned(native).ok()?;
    let offset = usize::from(native.starts_with('='));
    let body = &native[offset..];
    let syntax = crate::calc_print::CalcSyntax::for_decimal('.');
    let spans: Vec<_> = formula
        .ref_spans
        .iter()
        .zip(formula.expr.references())
        .map(|(span, r)| {
            let mut text = String::new();
            crate::calc_print::reference(&mut text, r, "");
            (span.start - offset..span.end - offset, text)
        })
        .collect();
    let mut next_ref = spans.iter().peekable();
    let mut out = String::with_capacity(body.len() + 8);
    let mut i = 0;
    while i < body.len() {
        if let Some((span, text)) = next_ref.peek()
            && span.start == i
        {
            out.push_str(text);
            i = span.end;
            next_ref.next();
            continue;
        }
        let rest = &body[i..];
        let c = rest.chars().next()?;
        if c == '"' {
            let mut end = 1;
            let b = rest.as_bytes();
            while end < b.len() {
                if b[end] == b'"' {
                    if b.get(end + 1) == Some(&b'"') {
                        end += 2;
                        continue;
                    }
                    end += 1;
                    break;
                }
                end += 1;
            }
            out.push_str(&rest[..end]);
            i += end;
        } else if c.is_alphabetic() || c == '_' {
            let len = rest
                .char_indices()
                .find(|(_, ch)| !(is_ident(*ch) || *ch == '.'))
                .map_or(rest.len(), |(n, _)| n);
            let word = &rest[..len];
            if rest[len..].starts_with('(') {
                // A function the application does not have is written in lowercase from an XLSX
                // (`casa()`, `_xlfn.casa()`).
                if crate::calc_print::has_function(word, known) {
                    out.push_str(&crate::calc_print::function_name(word, known));
                } else {
                    out.push_str(&word.to_lowercase());
                }
            } else if word.eq_ignore_ascii_case("TRUE") || word.eq_ignore_ascii_case("FALSE") {
                out.push_str(&word.to_ascii_uppercase());
                out.push_str("()");
            } else {
                out.push_str(word);
            }
            i += len;
        } else if c.is_ascii_digit()
            || (c == '.' && rest[1..].starts_with(|d: char| d.is_ascii_digit()))
        {
            let b = rest.as_bytes();
            let mut end = 0;
            while end < b.len() && (b[end].is_ascii_digit() || b[end] == b'.') {
                end += 1;
            }
            if matches!(b.get(end), Some(b'e' | b'E')) {
                let mut k = end + 1;
                if matches!(b.get(k), Some(b'+' | b'-')) {
                    k += 1;
                }
                if b.get(k).is_some_and(u8::is_ascii_digit) {
                    while b.get(k).is_some_and(u8::is_ascii_digit) {
                        k += 1;
                    }
                    end = k;
                }
            }
            match rest[..end].parse::<f64>() {
                Ok(v) => out.push_str(&crate::calc_print::number(v, &syntax)),
                Err(_) => out.push_str(&rest[..end]),
            }
            i += end;
        } else if c == '@' {
            i += 1;
        } else if c == '#'
            && rest.len() >= 5
            && rest.as_bytes()[..5].eq_ignore_ascii_case(b"#REF!")
            && !out.ends_with('!')
            && !rest[5..].starts_with('!')
        {
            // A bare `#REF!` is a name it does not know: `#ref!`.
            out.push_str("#ref!");
            i += 5;
        } else {
            out.push(c);
            i += c.len_utf8();
        }
    }
    Some(out)
}

/// An `OpenFormula` text in the notation `LibreOffice` shows (see [`native_to_calc`]).
fn openformula_to_calc(native: &str, known: &dyn Fn(&str) -> bool) -> String {
    let chars: Vec<char> = native.chars().collect();
    let mut out = String::with_capacity(native.len());
    let mut i = 0;
    let mut array = 0u32;
    while i < chars.len() {
        match chars[i] {
            '"' => {
                out.push('"');
                i += 1;
                while i < chars.len() {
                    out.push(chars[i]);
                    i += 1;
                    if chars[i - 1] == '"' {
                        if chars.get(i) == Some(&'"') {
                            out.push('"');
                            i += 1;
                        } else {
                            break;
                        }
                    }
                }
            }
            '[' => {
                // A reference: its text without the brackets, and without the dot of an empty
                // sheet (`.A1` at its start or after the `:`).
                i += 1;
                let mut quoted = false;
                let mut at_start = true;
                while i < chars.len() {
                    let c = chars[i];
                    if c == '\'' {
                        quoted = !quoted;
                    }
                    if quoted {
                        at_start = false;
                    } else {
                        if c == ']' {
                            i += 1;
                            break;
                        }
                        if c == '.' && at_start {
                            at_start = false;
                            i += 1;
                            continue;
                        }
                        at_start = c == ':';
                    }
                    out.push(c);
                    i += 1;
                }
            }
            '{' => {
                array += 1;
                out.push('{');
                i += 1;
            }
            '}' => {
                array = array.saturating_sub(1);
                out.push('}');
                i += 1;
            }
            ';' => {
                out.push(',');
                i += 1;
            }
            '|' if array > 0 => {
                out.push(';');
                i += 1;
            }
            c if c.is_alphabetic() || c == '_' => {
                // A name or a function: a function of a namespace (`COM.MICROSOFT.CONCAT`) is
                // shown without it when the application has it.
                let start = i;
                while i < chars.len() && (is_ident(chars[i]) || chars[i] == '.') {
                    i += 1;
                }
                let word: String = chars[start..i].iter().collect();
                let skip = crate::calc_print::namespace_len(&word);
                if chars.get(i) == Some(&'(')
                    && skip > 0
                    && known(&word[skip..].to_ascii_uppercase())
                {
                    out.push_str(&word[skip..]);
                } else {
                    out.push_str(&word);
                }
            }
            // The error `#REF!` a formula holds (`#ref!` is how `LibreOffice` writes it into an
            // ODS) is shown in capitals, unless it is the start of a name.
            '#' if chars.len() >= i + 5
                && chars[i..i + 5]
                    .iter()
                    .collect::<String>()
                    .eq_ignore_ascii_case("#REF!")
                && !chars.get(i + 5).is_some_and(|c| is_ident(*c)) =>
            {
                out.push_str("#REF!");
                i += 5;
            }
            c => {
                out.push(c);
                i += 1;
            }
        }
    }
    out
}

fn odf_sheet(name: &str) -> String {
    if crate::addr::needs_quotes(name) {
        format!("'{}'", name.replace('\'', "''"))
    } else {
        name.to_string()
    }
}

/// Reads `text` (a formula of `LibreOffice` without its `=`) and writes it in `notation`.
///
/// # Errors
///
/// When the text is not a formula, or holds a reference that spans several sheets.
pub fn from_calc(text: &str, notation: Notation) -> Result<String, FromCalcError> {
    let chars: Vec<char> = text.chars().collect();
    let r = Reader { s: &chars };
    let mut out = String::with_capacity(text.len() + 8);
    let mut i = 0;
    let mut array = 0u32;
    while i < chars.len() {
        let c = chars[i];
        match c {
            '"' => {
                out.push('"');
                i += 1;
                while i < chars.len() {
                    out.push(chars[i]);
                    i += 1;
                    if chars[i - 1] == '"' {
                        if chars.get(i) == Some(&'"') {
                            out.push('"');
                            i += 1;
                        } else {
                            break;
                        }
                    }
                }
            }
            '{' => {
                array += 1;
                out.push('{');
                i += 1;
            }
            '}' => {
                array = array.saturating_sub(1);
                out.push('}');
                i += 1;
            }
            ',' | ';' => {
                let in_array = array > 0;
                out.push(match (notation, in_array, c) {
                    (Notation::Excel, _, c) => {
                        if in_array {
                            c
                        } else {
                            ','
                        }
                    }
                    (Notation::OpenFormula, false, _) | (Notation::OpenFormula, true, ',') => ';',
                    (Notation::OpenFormula, true, _) => '|',
                });
                i += 1;
            }
            c if c == '$' || c == '\'' || c.is_alphabetic() || c.is_ascii_digit() => {
                if let Some((sheet, a, b, end)) = r.reference(i)? {
                    write_ref(&mut out, notation, sheet.as_ref(), &a, b.as_deref());
                    i = end;
                } else if is_ident(c) {
                    // A name, a function or a number: copied whole, so that what is inside it is not
                    // taken for a reference.
                    while i < chars.len()
                        && (is_ident(chars[i])
                            || (chars[i] == '.' && chars.get(i + 1).is_some_and(|n| is_ident(*n))))
                    {
                        out.push(chars[i]);
                        i += 1;
                    }
                } else {
                    out.push(c);
                    i += 1;
                }
            }
            c => {
                out.push(c);
                i += 1;
            }
        }
    }
    // The text must be a formula: checked in Excel's notation, whatever is written.
    let excel = if notation == Notation::Excel {
        out.clone()
    } else {
        from_calc(text, Notation::Excel)?
    };
    crate::parse::parse(&excel).map_err(|e| FromCalcError::NotAFormula(e.to_string()))?;
    Ok(out)
}

/// What `LibreOffice` makes of a text that was typed as a formula (see [`compile_calc`]).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Compiled {
    /// The text as the file keeps it: what `LibreOffice` regenerates from what it understood, in the
    /// notation of the file. Names it does not know are lowercased, and the parentheses left open at
    /// the end are closed.
    pub formula: String,
    /// The error `LibreOffice` gives the cell (`Err:511`, `525` is `#NAME?`), `None` when the
    /// formula is valid.
    pub error: Option<u16>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Tok {
    /// A number, text, reference, name or constant.
    Operand,
    /// `name(`.
    Function,
    Open,
    OpenArray,
    Close,
    CloseArray,
    /// `,` or `;`.
    Separator,
    Binary,
    /// `+` or `-` where an operand is expected.
    Prefix,
    /// `%`.
    Postfix,
    /// `:` between two parts of a range.
    Range,
}

/// What `LibreOffice` does with a text typed as a formula, probed in 26.8: it compiles the text and
/// keeps what it understood.
///
/// - The parentheses and braces left open at the end are closed (`SUM(1,2` is `SUM(1;2)`).
/// - A name it does not know (a function, a name, a reference to a sheet that is not there) makes
///   the cell `#NAME?` (525) and is written in lowercase.
/// - A text that is not a formula keeps its text, in an error cell: `Err:508` for a `)` that closes
///   nothing, `Err:509` for two operands with no operator between them (`A1 B1`, `1,2`), `Err:510`
///   for an operator where an operand is expected (`1+*2`), `Err:511` for a parenthesis with nothing
///   inside (`1+(`, `((`), `Err:520` for an operand that never comes (`A1+`, `A1:`), and `#NAME?`
///   for a text or a sheet name that is not closed.
/// - An error constant is its error (`#REF!` is 524).
///
/// `known_function`, `known_name` and `known_sheet` say whether the application, the workbook or the
/// workbook has such a function (uppercase, without prefix), defined name or sheet. The text is in
/// the notation of `en-US`; the result is written in `notation`.
#[must_use]
pub fn compile_calc(
    text: &str,
    notation: Notation,
    known_function: &dyn Fn(&str) -> bool,
    known_name: &dyn Fn(&str) -> bool,
    known_sheet: &dyn Fn(&str) -> bool,
) -> Compiled {
    let chars: Vec<char> = text.chars().collect();
    let r = Reader { s: &chars };
    let mut out = String::with_capacity(text.len() + 4);
    let mut toks: Vec<Tok> = Vec::new();
    let mut unknown = false;
    let mut unterminated = false;
    let mut constant_error: Option<u16> = None;
    let mut array = 0u32;
    let mut open: Vec<char> = Vec::new();
    let mut extra_close = false;
    let mut i = 0;
    let expect_operand = |toks: &[Tok]| {
        matches!(
            toks.last(),
            None | Some(
                Tok::Open
                    | Tok::OpenArray
                    | Tok::Function
                    | Tok::Separator
                    | Tok::Binary
                    | Tok::Prefix
                    | Tok::Range
            )
        )
    };
    while i < chars.len() {
        let c = chars[i];
        match c {
            c if c.is_whitespace() => {
                out.push(c);
                i += 1;
            }
            '"' => {
                let start = i;
                i += 1;
                let mut closed = false;
                while i < chars.len() {
                    if chars[i] == '"' {
                        if chars.get(i + 1) == Some(&'"') {
                            i += 2;
                            continue;
                        }
                        i += 1;
                        closed = true;
                        break;
                    }
                    i += 1;
                }
                out.extend(&chars[start..i]);
                unterminated |= !closed;
                toks.push(Tok::Operand);
            }
            '(' => {
                open.push(')');
                toks.push(Tok::Open);
                out.push(c);
                i += 1;
            }
            '{' => {
                open.push('}');
                array += 1;
                toks.push(Tok::OpenArray);
                out.push(c);
                i += 1;
            }
            ')' | '}' => {
                if open.last() == Some(&c) {
                    open.pop();
                    if c == '}' {
                        array = array.saturating_sub(1);
                    }
                } else {
                    extra_close = true;
                }
                toks.push(if c == ')' {
                    Tok::Close
                } else {
                    Tok::CloseArray
                });
                out.push(c);
                i += 1;
            }
            ',' | ';' => {
                let in_array = array > 0;
                out.push(match (notation, in_array, c) {
                    (Notation::Excel, true, c) => c,
                    (Notation::Excel, false, _) => ',',
                    (Notation::OpenFormula, false, _) | (Notation::OpenFormula, true, ',') => ';',
                    (Notation::OpenFormula, true, _) => '|',
                });
                toks.push(Tok::Separator);
                i += 1;
            }
            '+' | '-' => {
                toks.push(if expect_operand(&toks) {
                    Tok::Prefix
                } else {
                    Tok::Binary
                });
                out.push(c);
                i += 1;
            }
            '*' | '/' | '^' | '&' | '=' | '<' | '>' => {
                out.push(c);
                i += 1;
                if matches!((c, chars.get(i)), ('<', Some('>' | '=')) | ('>', Some('='))) {
                    out.push(chars[i]);
                    i += 1;
                }
                toks.push(Tok::Binary);
            }
            '%' => {
                toks.push(Tok::Postfix);
                out.push(c);
                i += 1;
            }
            ':' => {
                toks.push(Tok::Range);
                out.push(c);
                i += 1;
            }
            '#' => {
                let start = i;
                i += 1;
                while i < chars.len()
                    && (chars[i].is_alphanumeric() || matches!(chars[i], '/' | '!' | '?'))
                {
                    i += 1;
                }
                let word: String = chars[start..i].iter().collect();
                constant_error = Some(match word.to_ascii_uppercase().as_str() {
                    "#REF!" => 524,
                    "#NAME?" => 525,
                    "#DIV/0!" => 532,
                    "#VALUE!" => 519,
                    "#NUM!" => 503,
                    "#NULL!" => 521,
                    _ => 32767,
                });
                out.push_str(&word);
                toks.push(Tok::Operand);
            }
            c if c == '$' || c == '\'' || c.is_alphabetic() || c == '_' || c.is_ascii_digit() => {
                if c == '\'' && r.sheet(i).is_none() {
                    // An apostrophe that opens nothing: the text is not closed.
                    unterminated = true;
                    out.extend(&chars[i..]);
                    i = chars.len();
                    toks.push(Tok::Operand);
                    continue;
                }
                if let Ok(Some((sheet, a, b, end))) = r.reference(i) {
                    let known = sheet.as_ref().is_none_or(|(_, n)| known_sheet(n));
                    if known {
                        write_ref(&mut out, notation, sheet.as_ref(), &a, b.as_deref());
                    } else {
                        // A sheet that is not there: the text as typed, lowercased.
                        unknown = true;
                        let typed: String = chars[i..end].iter().collect();
                        out.push_str(&typed.to_lowercase());
                    }
                    toks.push(Tok::Operand);
                    i = end;
                } else if c.is_ascii_digit()
                    || (c == '.' && chars.get(i + 1).is_some_and(char::is_ascii_digit))
                {
                    let start = i;
                    while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
                        i += 1;
                    }
                    if matches!(chars.get(i), Some('e' | 'E')) {
                        let mut k = i + 1;
                        if matches!(chars.get(k), Some('+' | '-')) {
                            k += 1;
                        }
                        if chars.get(k).is_some_and(char::is_ascii_digit) {
                            while chars.get(k).is_some_and(char::is_ascii_digit) {
                                k += 1;
                            }
                            i = k;
                        }
                    }
                    out.extend(&chars[start..i]);
                    toks.push(Tok::Operand);
                } else if is_ident(c) {
                    let start = i;
                    while i < chars.len()
                        && (is_ident(chars[i])
                            || (chars[i] == '.' && chars.get(i + 1).is_some_and(|n| is_ident(*n))))
                    {
                        i += 1;
                    }
                    let word: String = chars[start..i].iter().collect();
                    let upper = word.to_ascii_uppercase();
                    if chars.get(i) == Some(&'(') {
                        let canonical = &upper[crate::calc_print::namespace_len(&word)..];
                        if known_function(canonical) {
                            out.push_str(&word);
                        } else {
                            unknown = true;
                            out.push_str(&word.to_lowercase());
                        }
                        toks.push(Tok::Function);
                        open.push(')');
                        out.push('(');
                        i += 1;
                    } else {
                        if upper == "TRUE" || upper == "FALSE" || known_name(&upper) {
                            out.push_str(&word);
                        } else {
                            unknown = true;
                            out.push_str(&word.to_lowercase());
                        }
                        toks.push(Tok::Operand);
                    }
                } else {
                    out.push(c);
                    i += 1;
                    toks.push(Tok::Operand);
                }
            }
            c => {
                out.push(c);
                i += 1;
                toks.push(Tok::Operand);
            }
        }
    }
    // What was left open is closed.
    if !extra_close {
        while let Some(closer) = open.pop() {
            out.push(closer);
            toks.push(if closer == ')' {
                Tok::Close
            } else {
                Tok::CloseArray
            });
        }
    }
    let error = if unterminated {
        Some(525)
    } else if let Some(code) = syntax_error(&toks) {
        Some(code)
    } else if unknown {
        Some(525)
    } else {
        constant_error
    };
    Compiled {
        formula: out,
        error,
    }
}

/// The error of a stream of tokens that is not an expression, `None` when it is one.
fn syntax_error(toks: &[Tok]) -> Option<u16> {
    let mut stack: Vec<Tok> = Vec::new();
    let mut operand = true; // an operand is expected
    for t in toks {
        match (*t, operand) {
            (Tok::Operand, true) => operand = false,
            (Tok::Operand | Tok::Function | Tok::Open | Tok::OpenArray, false) => return Some(509),
            (Tok::Function, true) => stack.push(Tok::Function),
            (Tok::Open, true) => stack.push(Tok::Open),
            (Tok::OpenArray, true) => stack.push(Tok::OpenArray),
            (Tok::Prefix, true) | (Tok::Postfix, false) => {}
            (Tok::Prefix | Tok::Binary | Tok::Range, false) => operand = true,
            (Tok::Binary | Tok::Postfix | Tok::Range, true) => return Some(510),
            (Tok::Separator, _) => match stack.last() {
                Some(Tok::Function | Tok::OpenArray) => operand = true,
                _ => return Some(509),
            },
            (Tok::Close, o) => match stack.pop() {
                Some(Tok::Function) => operand = false,
                Some(Tok::Open) => {
                    if o {
                        return Some(511);
                    }
                    operand = false;
                }
                _ => return Some(508),
            },
            (Tok::CloseArray, _) => match stack.pop() {
                Some(Tok::OpenArray) => operand = false,
                _ => return Some(508),
            },
        }
    }
    if operand && !toks.is_empty() {
        return Some(520);
    }
    if toks.is_empty() {
        return Some(511);
    }
    None
}

fn write_ref(
    out: &mut String,
    notation: Notation,
    sheet: Option<&(bool, String)>,
    a: &str,
    b: Option<&str>,
) {
    match notation {
        Notation::Excel => {
            if let Some((_, name)) = sheet {
                out.push_str(&sheet_text(name));
                out.push('!');
            }
            out.push_str(a);
            if let Some(b) = b {
                out.push(':');
                out.push_str(b);
            }
        }
        Notation::OpenFormula => {
            out.push('[');
            match sheet {
                Some((absolute, name)) => {
                    if *absolute {
                        out.push('$');
                    }
                    out.push_str(&odf_sheet(name));
                    out.push('.');
                }
                None => out.push('.'),
            }
            out.push_str(a);
            if let Some(b) = b {
                out.push_str(":.");
                out.push_str(b);
            }
            out.push(']');
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn excel(s: &str) -> String {
        from_calc(s, Notation::Excel).unwrap()
    }

    fn odf(s: &str) -> String {
        from_calc(s, Notation::OpenFormula).unwrap()
    }

    #[test]
    fn references_with_a_sheet_are_written_as_the_file_writes_them() {
        assert_eq!(excel("$'U Sheet'.A1+1"), "'U Sheet'!A1+1");
        assert_eq!(excel("SUM($T.A1:B3)"), "SUM(T!A1:B3)");
        assert_eq!(excel("SUM($T.A1:$T.B3)"), "SUM(T!A1:B3)");
        assert_eq!(excel("$'O''Brien'.A2"), "'O''Brien'!A2");
        assert_eq!(excel("A1+$B$2*C:C+3:4"), "A1+$B$2*C:C+3:4");
        assert_eq!(odf("$'U Sheet'.A1+1"), "[$'U Sheet'.A1]+1");
        assert_eq!(odf("SUM($T.A1:B3)"), "SUM([$T.A1:.B3])");
        assert_eq!(odf("A1+$B$2"), "[.A1]+[.$B$2]");
        assert_eq!(odf("SUM(A:A)"), "SUM([.A:.A])");
    }

    /// A bare `#REF!` is a name `LibreOffice` does not know in an XLSX (`#ref!`, `#NAME?`) and the
    /// error itself in an ODS, where it writes it `#ref!`; as part of a reference it is neither.
    #[test]
    fn a_bare_ref_error_is_told_from_one_inside_a_reference() {
        let excel = |f| has_bare_ref_error(f, Notation::Excel);
        let ods = |f| has_bare_ref_error(f, Notation::OpenFormula);
        for f in [
            "#REF!+5",
            "SUM(#REF!,1)",
            "IF(TRUE,#REF!,1)",
            "#ref!",
            "A5+#REF!",
            "#REF!A1",
        ] {
            assert!(excel(f), "{f}");
        }
        for f in [
            "Errors!#REF!",
            "#REF!!A1",
            "'Gone'!#REF!",
            "\"#REF!\"",
            "#N/A+1",
            "A1+B1",
        ] {
            assert!(!excel(f), "{f}");
        }
        for f in [
            "#ref!+5",
            "SUM(#ref!;1)",
            "[.A5]+#ref!",
            "IF(TRUE();#ref!;1)",
        ] {
            assert!(ods(f), "{f}");
        }
        // A reference that is gone, a name that starts with it, a quoted sheet and a text.
        for f in [
            "[#REF!]+1",
            "#ref!a1",
            "'Gone'!#ref!",
            "\"x\"&\"#ref!\"",
            "[.A1]",
        ] {
            assert!(!ods(f), "{f}");
        }
        let known = |name: &str| name != "CASA";
        assert_eq!(
            native_to_calc("#REF!+5", Notation::Excel, &known),
            "#ref!+5"
        );
        assert_eq!(
            native_to_calc("Errors!#REF!", Notation::Excel, &known),
            "Errors!#REF!"
        );
        assert_eq!(native_to_calc("#N/A+1", Notation::Excel, &known), "#N/A+1");
        assert_eq!(
            native_to_calc("#ref!+5", Notation::OpenFormula, &known),
            "#REF!+5"
        );
        assert_eq!(
            native_to_calc("[#REF!]+1", Notation::OpenFormula, &known),
            "#REF!+1"
        );
        assert_eq!(
            native_to_calc("#ref!a1", Notation::OpenFormula, &known),
            "#ref!a1"
        );
    }

    /// An XLSX shows the name of a function the application does not have in lowercase; an ODS keeps
    /// what the file says (it holds what `LibreOffice` itself wrote).
    #[test]
    fn an_unknown_function_of_an_xlsx_is_shown_in_lowercase() {
        let known = |name: &str| name != "CASA";
        let xlsx = |f| native_to_calc(f, Notation::Excel, &known);
        assert_eq!(xlsx("CASA()"), "casa()");
        assert_eq!(xlsx("_xlfn.CASA()+SUM(1,2)"), "_xlfn.casa()+SUM(1,2)");
        assert_eq!(xlsx("casa(1,2)+1"), "casa(1,2)+1");
        assert_eq!(xlsx("sum(1,2)"), "SUM(1,2)");
        assert_eq!(
            native_to_calc("CASA()", Notation::OpenFormula, &known),
            "CASA()"
        );
    }

    #[test]
    fn a_native_formula_is_shown_as_the_application_shows_it() {
        let known = |name: &str| name != "FOO";
        let calc = |s: &str, n| native_to_calc(s, n, &known);
        // An ODS keeps the sheets that are relative and the ones that are absolute.
        assert_eq!(calc("[T.A1]*2", Notation::OpenFormula), "T.A1*2");
        assert_eq!(calc("[$T.A1]*2", Notation::OpenFormula), "$T.A1*2");
        assert_eq!(calc("[.A1]+[.$B$2]", Notation::OpenFormula), "A1+$B$2");
        assert_eq!(
            calc("SUM([T.A1:T.B3])", Notation::OpenFormula),
            "SUM(T.A1:T.B3)"
        );
        assert_eq!(
            calc("SUM([T.A1:.A3])", Notation::OpenFormula),
            "SUM(T.A1:A3)"
        );
        assert_eq!(
            calc("['Data-1'.A1:'Data-1'.A3]", Notation::OpenFormula),
            "'Data-1'.A1:'Data-1'.A3"
        );
        assert_eq!(calc("SUM([.A:.A])", Notation::OpenFormula), "SUM(A:A)");
        assert_eq!(
            calc("IF([.A1]>0;\"a;[.B1]\";{1;2|3;4})", Notation::OpenFormula),
            "IF(A1>0,\"a;[.B1]\",{1,2;3,4})"
        );
        assert_eq!(calc("foo()", Notation::OpenFormula), "foo()");
        assert_eq!(
            calc("COM.MICROSOFT.CONCAT([.A1];\"x\")", Notation::OpenFormula),
            "CONCAT(A1,\"x\")"
        );
        assert_eq!(
            calc("LOG10([.A1])+TRUE()", Notation::OpenFormula),
            "LOG10(A1)+TRUE()"
        );
        // An XLSX keeps the sheet it was written with, the own sheet included.
        assert_eq!(calc("T!A1*2", Notation::Excel), "$T.A1*2");
        assert_eq!(calc("A1*2", Notation::Excel), "A1*2");
        assert_eq!(calc("'Data-1'!A1:A3", Notation::Excel), "$'Data-1'.A1:A3");
        assert_eq!(calc("foo()+sum(A1)", Notation::Excel), "foo()+SUM(A1)");
        // The spacing is the one the file has; the numbers and the constants are the application's.
        assert_eq!(
            calc("IF(I7=\"\", \"Y\",\"N\")", Notation::Excel),
            "IF(I7=\"\", \"Y\",\"N\")"
        );
        assert_eq!(
            calc("VLOOKUP(1.50, A1:B2 ,2,FALSE)+@C1", Notation::Excel),
            "VLOOKUP(1.5, A1:B2 ,2,FALSE())+C1"
        );
        assert_eq!(
            calc("_xlfn.CONCAT(A1,\"1.50\")", Notation::Excel),
            "CONCAT(A1,\"1.50\")"
        );
    }

    fn compiled(text: &str) -> (String, Option<u16>) {
        // The functions the application has, the names and sheets of the workbook: the ones of these
        // cases.
        let c = compile_calc(
            text,
            Notation::Excel,
            &|f| matches!(f, "SUM" | "IF" | "LEN" | "CONCAT"),
            &|_| false,
            &|s| s.eq_ignore_ascii_case("T"),
        );
        (c.formula, c.error)
    }

    // Each case is what `LibreOffice` 26.8 made of the text typed into a formula cell (the cell is
    // `=SUM(1;2)` and the text replaces `SUM(1,2)`; see `tests/assets/replace/make_fixtures.py`).
    #[test]
    fn a_text_typed_as_a_formula_is_what_libreoffice_makes_of_it() {
        let ok = |s: &str| compiled(s);
        assert_eq!(ok("SUM(1,2"), ("SUM(1,2)".into(), None));
        assert_eq!(ok("IF(1,2"), ("IF(1,2)".into(), None));
        assert_eq!(ok("LEN(\"a\""), ("LEN(\"a\")".into(), None));
        assert_eq!(ok("{1,2"), ("{1,2}".into(), None));
        assert_eq!(ok("1+("), ("1+()".into(), Some(511)));
        assert_eq!(ok("(("), ("(())".into(), Some(511)));
        assert_eq!(ok("SUM(1,2)*(("), ("SUM(1,2)*(())".into(), Some(511)));
        assert_eq!(ok("SUM(1,2))"), ("SUM(1,2))".into(), Some(508)));
        assert_eq!(ok("A1)"), ("A1)".into(), Some(508)));
        assert_eq!(ok("A1+"), ("A1+".into(), Some(520)));
        assert_eq!(ok("A1:"), ("A1:".into(), Some(520)));
        assert_eq!(ok("SUM(A1:A3)*"), ("SUM(A1:A3)*".into(), Some(520)));
        assert_eq!(ok("1/0+"), ("1/0+".into(), Some(520)));
        assert_eq!(ok("1+*2"), ("1+*2".into(), Some(510)));
        assert_eq!(ok("A1 B1"), ("A1 B1".into(), Some(509)));
        assert_eq!(ok("1 2"), ("1 2".into(), Some(509)));
        assert_eq!(ok("1,2"), ("1,2".into(), Some(509)));
        // What it does not know is `#NAME?`, in lowercase.
        assert_eq!(ok("foo(1)"), ("foo(1)".into(), Some(525)));
        assert_eq!(ok("Foo"), ("foo".into(), Some(525)));
        assert_eq!(ok("A1+B"), ("A1+b".into(), Some(525)));
        assert_eq!(ok("SUMX(1,2)*2"), ("sumx(1,2)*2".into(), Some(525)));
        assert_eq!(ok("$Nope.A1"), ("$nope.a1".into(), Some(525)));
        assert_eq!(ok("A1+Sheet9.A1"), ("A1+sheet9.a1".into(), Some(525)));
        assert_eq!(ok("\"abc"), ("\"abc".into(), Some(525)));
        assert_eq!(ok("'abc"), ("'abc".into(), Some(525)));
        assert_eq!(ok("#REF!+1"), ("#REF!+1".into(), Some(524)));
        // And what it knows is a formula, written in the notation of the file.
        assert_eq!(ok("$T.A1*2"), ("T!A1*2".into(), None));
    }

    #[test]
    fn separators_and_arrays_follow_the_notation() {
        assert_eq!(
            excel("IF(A1>0;\"a,b\";{1,2;3,4})"),
            "IF(A1>0,\"a,b\",{1,2;3,4})"
        );
        assert_eq!(
            odf("IF(A1>0,\"a;b\",{1,2;3,4})"),
            "IF([.A1]>0;\"a;b\";{1;2|3;4})"
        );
    }

    #[test]
    fn text_names_and_functions_are_not_references() {
        assert_eq!(excel("\"A1 $T.A1\"&A1"), "\"A1 $T.A1\"&A1");
        assert_eq!(excel("LOG10(2)+TRUE()"), "LOG10(2)+TRUE()");
        assert_eq!(odf("LOG10(A1)"), "LOG10([.A1])");
        assert_eq!(
            excel("COM.MICROSOFT.CONCAT(A1,B1)"),
            "COM.MICROSOFT.CONCAT(A1,B1)"
        );
        assert_eq!(excel("1.5+2E3"), "1.5+2E3");
    }

    #[test]
    fn what_is_not_a_formula_is_refused() {
        assert!(from_calc("((*2", Notation::Excel).is_err());
        assert!(from_calc("$A.A1:$B.A1", Notation::Excel).is_err());
    }

    #[test]
    fn it_is_the_inverse_of_the_printer() {
        use crate::calc_print::{CalcSyntax, to_calc};
        for src in [
            "'My Sheet'!A1+F160!B2:C3+S!A1",
            "SUM(A1:A3,2)*IF(A1>1,\"a\"\"b\",{1,2;3,4})",
            "$A$1+B$2-A:A",
        ] {
            let calc = to_calc(
                &crate::parse(src).unwrap(),
                "S",
                &CalcSyntax::for_decimal('.'),
            );
            let back = excel(&calc);
            let again = to_calc(
                &crate::parse(&back).unwrap(),
                "S",
                &CalcSyntax::for_decimal('.'),
            );
            assert_eq!(calc, again, "{src}");
        }
    }
}

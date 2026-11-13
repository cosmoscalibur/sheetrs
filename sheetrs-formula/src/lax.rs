//! What `LibreOffice` makes of a formula text that the parser refuses, for the texts that have an
//! array in them or are left open at the end (probed in 26.8 with `sheetrs-compiler/scripts/lo_probe.py`).
//!
//! A formula in a file is compiled by `LibreOffice` when it opens it, and what it cannot read is
//! not a failure of the load: the cell holds an error. Two things matter here:
//!
//! - The delimiters left open at the end are closed (`{1,2` is `{1,2}`, `SUM({1,2},{3` is
//!   `SUM({1,2},{3})`), as long as what is inside is complete: an array that ends after a separator
//!   (`{1,`), with nothing in it (`{`) or with something in it that an array does not take
//!   (`{1 2`, `{1,'y'`) is `Err:512`.
//! - What an array holds is read when its text is read, before the rest of the formula: a
//!   `{1 2}` or a `{"x",'y'}` is `Err:539` whatever follows it, an array in an array is `Err:533`, and
//!   only then the text around it is checked, as it is for any formula (`Err:508` for a closing
//!   delimiter that closes nothing, `Err:509` for two operands with nothing between them, `{1,2}{3}`
//!   or `{1,2},3`; see [`crate::calc_read::compile_calc`]).

use crate::calc_read::{Notation, compile_calc};

/// What the formula text is.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Lax {
    /// The text with the delimiters it left open closed: it is a formula.
    Closed(String),
    /// The text is not a formula: the error the cell gets (`Err:512`, `Err:539`…).
    Error(u16),
}

/// How the scan of an array written in the text ended.
enum Array {
    /// Closed at the byte index (after the `}`); whether something in it is not an array element.
    Closed(usize, bool),
    /// The text ended inside it: whether what is in it is not complete or not an element.
    Open(bool),
    /// Another array opens inside it.
    Nested,
    /// A `)` where an element or a separator goes.
    Paren,
}

/// `LibreOffice`'s reading of `src`, when the parser cannot read it and `LibreOffice` can say what
/// it is. `None` for a text that is not one of those (the parser's refusal stands).
#[must_use]
pub fn lax(src: &str) -> Option<Lax> {
    let body = src.strip_prefix('=').unwrap_or(src);
    // What is wrong in a text with no array is said only for what no formula can mean: a space between
    // two references is the intersection operator in a file, an error only when it is typed.
    let arrays = body.contains(['{', '}']);
    let chars: Vec<char> = body.chars().collect();
    let mut out = String::with_capacity(body.len());
    let mut stack: Vec<char> = Vec::new();
    let mut first: Option<u16> = None;
    let mut i = 0;
    while i < chars.len() {
        match chars[i] {
            '"' => {
                let start = i;
                i += 1;
                loop {
                    match chars.get(i) {
                        None => return None,
                        Some('"') if chars.get(i + 1) == Some(&'"') => i += 2,
                        Some('"') => {
                            i += 1;
                            break;
                        }
                        Some(_) => i += 1,
                    }
                }
                out.extend(&chars[start..i]);
            }
            '{' => match scan_array(&chars, i) {
                Array::Closed(end, bad) => {
                    if bad {
                        first.get_or_insert(539);
                    }
                    // The array is one operand for what is around it.
                    out.push_str(" 0 ");
                    i = end;
                }
                Array::Open(bad) => {
                    // The text ended inside it: it closes if what is in it is complete.
                    if bad {
                        first.get_or_insert(512);
                    } else {
                        stack.push('}');
                        out.push_str(" 0 ");
                    }
                    i = chars.len();
                }
                Array::Nested => {
                    first.get_or_insert(533);
                    break;
                }
                Array::Paren => {
                    first.get_or_insert(if stack.is_empty() { 508 } else { 512 });
                    break;
                }
            },
            '(' => {
                stack.push(')');
                out.push('(');
                i += 1;
            }
            ')' => {
                if stack.last() == Some(&')') {
                    stack.pop();
                } else if stack.is_empty() {
                    first.get_or_insert(508);
                }
                out.push(')');
                i += 1;
            }
            '}' => {
                // Closes nothing: outside an array a `}` is not part of any formula.
                first.get_or_insert(508);
                out.push('}');
                i += 1;
            }
            c => {
                out.push(c);
                i += 1;
            }
        }
    }
    if let Some(code) = first {
        return Some(Lax::Error(code));
    }
    let last = chars.iter().rev().find(|c| !c.is_whitespace());
    if !stack.is_empty() && matches!(last, Some(',' | ';')) {
        return Some(Lax::Error(511));
    }
    let known = |_: &str| true;
    if let Some(code) = compile_calc(&out, Notation::Excel, &known, &known, &known).error
        && (matches!(code, 508 | 510 | 511 | 520) || (arrays && code == 509))
    {
        return Some(Lax::Error(code));
    }
    if stack.is_empty() {
        return None;
    }
    let mut closed = body.to_string();
    closed.extend(stack.iter().rev());
    Some(Lax::Closed(closed))
}

/// The scan of the array that opens at `start` (a `{`): the elements an array takes are numbers
/// (with a sign), texts, `TRUE` and `FALSE` and errors, between separators.
fn scan_array(chars: &[char], start: usize) -> Array {
    let mut j = start + 1;
    let mut bad = false;
    // An element is expected: at the start and after a separator.
    let mut expect = true;
    // The last thing was a separator or the opening: the array is not complete if the text ends.
    let mut dangling = true;
    loop {
        while chars.get(j).is_some_and(|c| c.is_whitespace()) {
            j += 1;
        }
        let Some(&c) = chars.get(j) else {
            return Array::Open(bad || dangling);
        };
        match c {
            '}' => return Array::Closed(j + 1, bad),
            '{' => return Array::Nested,
            ')' => return Array::Paren,
            ',' | ';' => {
                expect = true;
                dangling = true;
                j += 1;
            }
            _ => {
                if !expect {
                    bad = true;
                }
                expect = false;
                dangling = false;
                match c {
                    '"' => {
                        j += 1;
                        loop {
                            match chars.get(j) {
                                None => return Array::Open(true),
                                Some('"') if chars.get(j + 1) == Some(&'"') => j += 2,
                                Some('"') => {
                                    j += 1;
                                    break;
                                }
                                Some(_) => j += 1,
                            }
                        }
                    }
                    '+' | '-' | '0'..='9' | '.' => {
                        if matches!(c, '+' | '-') {
                            j += 1;
                        }
                        let from = j;
                        while chars
                            .get(j)
                            .is_some_and(|c| c.is_ascii_digit() || *c == '.')
                        {
                            j += 1;
                        }
                        if matches!(chars.get(j), Some('e' | 'E')) {
                            let mut k = j + 1;
                            if matches!(chars.get(k), Some('+' | '-')) {
                                k += 1;
                            }
                            if chars.get(k).is_some_and(char::is_ascii_digit) {
                                while chars.get(k).is_some_and(char::is_ascii_digit) {
                                    k += 1;
                                }
                                j = k;
                            }
                        }
                        if j == from {
                            bad = true;
                        }
                    }
                    '#' => {
                        // An error: its spelling is for the compiler of the formula to check.
                        while chars.get(j).is_some_and(|c| {
                            !matches!(c, ',' | ';' | '}' | ')') && !c.is_whitespace()
                        }) {
                            j += 1;
                        }
                    }
                    c if c.is_alphabetic() => {
                        let from = j;
                        while chars
                            .get(j)
                            .is_some_and(|c| c.is_alphanumeric() || *c == '_')
                        {
                            j += 1;
                        }
                        let word: String = chars[from..j].iter().collect();
                        if !(word.eq_ignore_ascii_case("TRUE")
                            || word.eq_ignore_ascii_case("FALSE"))
                        {
                            bad = true;
                        }
                    }
                    _ => {
                        // Anything else (`'y'`, `(1)`, `@`) is not an element: skip to the next
                        // separator.
                        bad = true;
                        while let Some(&d) = chars.get(j) {
                            if matches!(d, ',' | ';' | '}' | ')') {
                                break;
                            }
                            j += 1;
                            if d == '\'' {
                                while chars.get(j).is_some_and(|d| *d != '\'') {
                                    j += 1;
                                }
                                j += usize::from(j < chars.len());
                            }
                        }
                    }
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn err(src: &str) -> Option<u16> {
        match lax(src)? {
            Lax::Error(c) => Some(c),
            Lax::Closed(_) => None,
        }
    }

    fn closed(src: &str) -> Option<String> {
        match lax(src)? {
            Lax::Closed(t) => Some(t),
            Lax::Error(_) => None,
        }
    }

    #[test]
    fn what_libreoffice_gives_the_texts_it_cannot_read() {
        // Probed in LibreOffice 26.8 (`sheetrs-compiler/scripts/lo_probe.py`).
        for (src, code) in [
            ("{1,2}{3}", 509),
            ("{1,2}}", 508),
            ("{1,2}}}", 508),
            ("{1 2}", 539),
            ("{\"x\",'y'}", 539),
            ("{\"a\" \"b\"}", 539),
            ("{1,2}{3,4}", 509),
            ("{1,2};{3,4}", 509),
            ("{1,2}(3)", 509),
            ("{1}{2}", 509),
            ("{{1}}", 533),
            ("{{", 533),
            ("{1,{2", 533),
            ("{1,2}  {3}", 509),
            ("{1,2}\"a\"", 509),
            ("{1,2}1", 509),
            ("{1,2}x", 509),
            ("{1,2}{", 512),
            ("}", 508),
            ("{", 512),
            ("{1,", 512),
            ("{1;", 512),
            ("{1,2;", 512),
            ("{\"a\",", 512),
            ("{1,2}+{", 512),
            ("{1,2}+{3,", 512),
            ("{1 2", 512),
            ("{1,'y'", 512),
            ("SUM({1,2}{3})", 509),
            ("SUM({1 2})", 539),
            ("SUM({1,2}}", 508),
            ("{1,2)", 508),
            ("({1,2)", 512),
            ("{1,2}{3 4}", 539),
            ("{1,2},3", 509),
            ("{1,2},{3,4}", 509),
            ("{1,2}}x", 508),
            ("{1,2}}{", 508),
            ("SUM({1,2},{", 512),
            ("SUM({1,2},{3,", 512),
            ("{1,2}{3,4}}", 508),
            ("1}", 508),
            ("SUM({1,", 512),
            ("({", 512),
            ("SUM(1,", 511),
            ("(", 511),
            ("1+", 520),
            ("1)", 508),
            ("(1+2))", 508),
            ("SUM(1,2))", 508),
        ] {
            assert_eq!(err(src), Some(code), "{src}");
        }
    }

    #[test]
    fn what_is_left_open_is_closed() {
        for (src, text) in [
            ("{1,2", "{1,2}"),
            ("{1,2,3", "{1,2,3}"),
            ("{1;2", "{1;2}"),
            ("{1,2;3,4", "{1,2;3,4}"),
            ("{\"a\"", "{\"a\"}"),
            ("{\"a\",\"b\"", "{\"a\",\"b\"}"),
            ("{1,2}+{3", "{1,2}+{3}"),
            ("SUM({1,2", "SUM({1,2})"),
            ("SUM({1,2},{3", "SUM({1,2},{3})"),
            ("{1,2}{3,4", "{1,2}{3,4}"),
            ("SUM(1,2", "SUM(1,2)"),
            ("(1+2", "(1+2)"),
            ("SUM(", "SUM()"),
            ("1+(2", "1+(2)"),
        ] {
            // `{1,2}{3,4` is `Err:509` once closed: the closing is not what makes it a formula.
            if src == "{1,2}{3,4" {
                assert_eq!(err(src), Some(509), "{src}");
                continue;
            }
            assert_eq!(closed(src).as_deref(), Some(text), "{src}");
        }
    }

    #[test]
    fn a_text_that_is_not_one_of_those_is_left_to_the_parser() {
        assert_eq!(lax("SUM(1,2)"), None);
        assert_eq!(lax("A1 B1"), None);
        assert_eq!(lax("Data!A1:Wide!A1"), None);
        assert_eq!(lax("{1,2}"), None);
        assert_eq!(lax("Sheet1:Sheet3!A1"), None);
    }
}

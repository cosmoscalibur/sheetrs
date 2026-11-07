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

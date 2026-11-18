//! Reading of cell contents the way `LibreOffice` 26.8 loads them, including its
//! behaviour for text that is not what the cell type announces.

use std::borrow::Cow;

use sheetrs_core::{ErrorKind, MAX_COLS, MAX_ROWS};
use sheetrs_io::{XML_WHITESPACE, decimal_prefix, integer_prefix_skipping, is_wide_whitespace};

/// The whitespace skipped before a stored number.
const NUMBER_WHITESPACE: [char; 2] = [' ', '\t'];

/// The number a stored text stands for.
///
/// After optional leading spaces and tabs and a sign, the longest prefix that is a decimal
/// number counts; text without such a prefix is zero. `NaN` (unsigned), `1.#NAN`, `INF` and
/// `1.#INF` are the words for not-a-number and infinity.
pub(crate) fn number(text: &str) -> f64 {
    let rest = text.trim_start_matches(NUMBER_WHITESPACE);
    let (negative, unsigned, signed) = match rest.strip_prefix('-') {
        Some(after) => (true, after, true),
        None => match rest.strip_prefix('+') {
            Some(after) => (false, after, true),
            None => (false, rest, false),
        },
    };
    if (!signed && unsigned.starts_with("NaN")) || unsigned.starts_with("1.#NAN") {
        return f64::NAN;
    }
    if unsigned.starts_with("INF") || unsigned.starts_with("1.#INF") {
        return if negative {
            f64::NEG_INFINITY
        } else {
            f64::INFINITY
        };
    }
    let magnitude = decimal_prefix(unsigned).map_or(0.0, |prefix| prefix.parse().unwrap_or(0.0));
    if negative { -magnitude } else { magnitude }
}

/// A stored boolean: `true` in any case, or any number other than zero.
pub(crate) fn boolean(text: &str) -> bool {
    if text.eq_ignore_ascii_case("true") {
        return true;
    }
    let value = number(text);
    value.is_nan() || value.abs() > 0.0
}

/// The numbers between brackets that the expression uses to refer to other documents, such as
/// the 1 of `[1]Sheet!A1`, outside quoted text.
pub(crate) fn external_indices(expression: &str) -> Vec<usize> {
    let mut found = Vec::new();
    let mut quote: Option<char> = None;
    let mut digits: Option<String> = None;
    for character in expression.chars() {
        if let Some(open) = quote {
            if character == open {
                quote = None;
            }
        } else if let Some(number) = digits.as_mut() {
            if character.is_ascii_digit() {
                number.push(character);
            } else {
                if character == ']'
                    && let Ok(index) = number.parse()
                {
                    found.push(index);
                }
                digits = None;
            }
        } else if character == '"' || character == '\'' {
            quote = Some(character);
        } else if character == '[' {
            digits = Some(String::new());
        }
    }
    found
}

/// Whether `expression` is nothing but a reference to a cell or a range of another document,
/// such as `[1]Sheet!$A$1:$B$2`.
pub(crate) fn is_external_reference(expression: &str) -> bool {
    let Some(rest) = expression.strip_prefix('[') else {
        return false;
    };
    let Some((index, rest)) = rest.split_once(']') else {
        return false;
    };
    if index.is_empty() || !index.bytes().all(|byte| byte.is_ascii_digit()) {
        return false;
    }
    let Some((_, cells)) = rest.rsplit_once('!') else {
        return false;
    };
    let cell = |text: &str| {
        let text = text.replace('$', "");
        let letters = text.bytes().take_while(u8::is_ascii_alphabetic).count();
        letters > 0
            && letters <= 3
            && text.len() > letters
            && text[letters..].bytes().all(|byte| byte.is_ascii_digit())
    };
    cells.split(':').count() <= 2 && cells.split(':').all(cell)
}

/// Whether a flag attribute of the workbook part, such as `date1904` or `hidden`, is on: the
/// words `true`, `on` and `t`, or a number other than zero.
pub(crate) fn workbook_flag(text: &str) -> bool {
    if matches!(text, "true" | "on" | "t") {
        return true;
    }
    let rest = text.trim_start_matches(XML_WHITESPACE);
    let unsigned = rest.strip_prefix(['-', '+']).unwrap_or(rest);
    unsigned
        .bytes()
        .take_while(u8::is_ascii_digit)
        .any(|digit| digit != b'0')
}

/// The index a stored shared-string reference stands for, or `None` when it is negative.
pub(crate) fn shared_string_index(text: &str) -> Option<usize> {
    usize::try_from(integer_prefix_skipping(text, is_wide_whitespace)).ok()
}

/// The error a cell of type `e` without a formula holds, and the formula `LibreOffice`
/// makes of its text when that text is the name of an error.
///
/// The six classic errors are matched ignoring case and surrounding whitespace; other text
/// is read as the formula it would be, so most of it is `#NAME?` and has no formula here.
pub(crate) fn constant_error(text: &str) -> (ErrorKind, Option<&'static str>) {
    const CLASSIC: [(&str, ErrorKind); 6] = [
        ("#NULL!", ErrorKind::Null),
        ("#DIV/0!", ErrorKind::DivZero),
        ("#VALUE!", ErrorKind::Value),
        ("#NUM!", ErrorKind::Num),
        ("#N/A", ErrorKind::Na),
        ("#NAME?", ErrorKind::Name),
    ];
    let upper = text.trim_matches(XML_WHITESPACE).to_ascii_uppercase();
    if upper.is_empty() {
        return (ErrorKind::Other("Err:520".into()), None);
    }
    for (name, kind) in CLASSIC {
        if upper == name {
            return (kind, Some(name));
        }
        if upper.starts_with(name) {
            return (ErrorKind::Other("Err:509".into()), None);
        }
    }
    match upper.as_str() {
        "#SPILL!" => (ErrorKind::Other("#SPILL!".into()), Some("#SPILL!")),
        "#GETTING_DATA" => (ErrorKind::Other("Err:509".into()), None),
        _ => (ErrorKind::Name, None),
    }
}

/// The functions, as an XLSX file spells them, whose cells `LibreOffice` calculates again on load
/// even when they cache a number.
const RECALCULATED_FUNCTIONS: [&str; 12] = [
    "NOW",
    "TODAY",
    "RAND",
    "RANDBETWEEN",
    "OFFSET",
    "INDIRECT",
    "INFO",
    "CELL",
    "HYPERLINK",
    "_xlfn.RANDARRAY",
    "_xlfn.FILTERXML",
    "_xlfn.FORMULATEXT",
];

/// Whether `formula` calls a function that `LibreOffice` calculates again when it loads the cell.
pub(crate) fn recalculated_on_load(formula: &str) -> bool {
    sheetrs_io::calls_any(formula, &RECALCULATED_FUNCTIONS)
}

/// The text of a shared or inline string with its line breaks normalized: a carriage return
/// and a line feed in either order are one line feed.
pub(crate) fn normalize_line_breaks(text: &str) -> Cow<'_, str> {
    if !text.contains('\r') {
        return Cow::Borrowed(text);
    }
    let mut normalized = String::with_capacity(text.len());
    let mut characters = text.chars().peekable();
    while let Some(character) = characters.next() {
        let partner = match character {
            '\r' => Some('\n'),
            '\n' => Some('\r'),
            _ => None,
        };
        if partner.is_some() && characters.peek() == partner.as_ref() {
            characters.next();
            normalized.push('\n');
        } else {
            normalized.push(character);
        }
    }
    Cow::Owned(normalized)
}

/// The cells a `ref` attribute covers, as `LibreOffice` reads it.
pub(crate) struct Area {
    first_col: u32,
    last_col: u32,
    first_row: u64,
    last_row: u64,
    /// Whether the end of the range is a column alone or an incomplete row, so that its last
    /// row is the end of the sheet, which an array formula does not follow.
    pub(crate) open_ended: bool,
}

impl Area {
    /// Whether the cell at zero-based `row` and `col` lies inside the area.
    pub(crate) fn contains(&self, row: u32, col: u32) -> bool {
        let (row, col) = (u64::from(row) + 1, col + 1);
        (self.first_col..=self.last_col).contains(&col)
            && (self.first_row..=self.last_row).contains(&row)
    }

    /// Whether the cell at zero-based `row` and `col` is the first row and column of the area.
    pub(crate) fn starts_at(&self, row: u32, col: u32) -> bool {
        self.first_row == u64::from(row) + 1 && self.first_col == col + 1
    }

    /// The last row and column of the area, zero-based and cut at `rows` and `cols`.
    pub(crate) fn end_within(&self, rows: u32, cols: u32) -> (u32, u32) {
        let row = self.last_row.min(u64::from(rows));
        (
            u32::try_from(row - 1).expect("the cut fits 32 bits"),
            self.last_col.min(cols) - 1,
        )
    }
}

/// One side of a range, with its 1-based numbers.
enum Corner {
    Cell {
        col: u32,
        row: u64,
    },
    Column(u32),
    Row(u64),
    /// The end of a range whose row `LibreOffice` does not take as a row: it covers the column
    /// down to the last row and ignores what follows the column.
    Open {
        col: u32,
    },
}

/// The 1-based number of a column written in letters, if the sheet has it.
fn column_number(letters: &str) -> Option<u32> {
    letters.bytes().try_fold(0_u32, |col, byte| {
        let col = col * 26 + u32::from(byte.to_ascii_uppercase() - b'A' + 1);
        (col <= MAX_COLS).then_some(col)
    })
}

/// Parses one side of a range. The end of a range is lenient about a row that is zero,
/// negative or beyond the sheet; a row inside the sheet must be followed by nothing.
fn corner(text: &str, end: bool) -> Option<Corner> {
    let text = text.strip_prefix('$').unwrap_or(text);
    let letters = text
        .find(|character: char| !character.is_ascii_alphabetic())
        .unwrap_or(text.len());
    let (letters, rest) = text.split_at(letters);
    if letters.is_empty() {
        let row = rest.parse::<u64>().ok().filter(|row| *row >= 1)?;
        return rest
            .bytes()
            .all(|byte| byte.is_ascii_digit())
            .then_some(Corner::Row(row));
    }
    let col = column_number(letters)?;
    if rest.is_empty() {
        return Some(Corner::Column(col));
    }
    let rest = rest.strip_prefix('$').unwrap_or(rest);
    if rest.is_empty() {
        return end.then_some(Corner::Open { col });
    }
    let (negative, body) = match rest.as_bytes().first() {
        Some(b'+') => (false, &rest[1..]),
        Some(b'-') => (true, &rest[1..]),
        Some(byte) if byte.is_ascii_digit() => (false, rest),
        _ => return None,
    };
    let digits = body.bytes().take_while(u8::is_ascii_digit).count();
    let row = body[..digits].bytes().fold(0_u64, |row, byte| {
        row.saturating_mul(10)
            .saturating_add(u64::from(byte - b'0'))
    });
    let in_sheet = !negative && (1..=u64::from(MAX_ROWS)).contains(&row);
    if in_sheet {
        (digits == body.len()).then_some(Corner::Cell { col, row })
    } else {
        end.then_some(Corner::Open { col })
    }
}

/// Reads a `ref` attribute: a cell, or a range of cells, of whole columns or of whole rows,
/// in either order and followed by spaces. `None` when `LibreOffice` rejects it.
pub(crate) fn area(text: &str) -> Option<Area> {
    let text = text.trim_end_matches(' ');
    let Some((first, last)) = text.split_once(':') else {
        return match corner(text, false)? {
            Corner::Cell { col, row } => Some(Area {
                first_col: col,
                last_col: col,
                first_row: row,
                last_row: row,
                open_ended: false,
            }),
            _ => None,
        };
    };
    Some(match (corner(first, false)?, corner(last, true)?) {
        (
            Corner::Cell { col, row },
            Corner::Cell {
                col: col2,
                row: row2,
            },
        ) => Area {
            first_col: col.min(col2),
            last_col: col.max(col2),
            first_row: row.min(row2),
            last_row: row.max(row2),
            open_ended: false,
        },
        (Corner::Cell { col, row }, Corner::Column(col2) | Corner::Open { col: col2 }) => Area {
            first_col: col.min(col2),
            last_col: col.max(col2),
            first_row: row,
            last_row: u64::MAX,
            open_ended: true,
        },
        (Corner::Column(col), Corner::Column(col2)) => Area {
            first_col: col.min(col2),
            last_col: col.max(col2),
            first_row: 1,
            last_row: u64::MAX,
            open_ended: false,
        },
        (Corner::Row(row), Corner::Row(row2)) => Area {
            first_col: 1,
            last_col: MAX_COLS,
            first_row: row.min(row2),
            last_row: row.max(row2),
            open_ended: false,
        },
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn functions_that_are_calculated_again_are_found_as_calls() {
        for formula in [
            "NOW()",
            "now ()",
            "1+TODAY()",
            "IF(1,Rand(),2)",
            "SUM(A1,OFFSET(B1,0,0))",
            "_xlfn.RANDARRAY()",
            "'my sheet(1)'!B1+INFO(\"system\")",
            "\"it\"\"s\"&CELL(\"row\")",
        ] {
            assert!(recalculated_on_load(formula), "{formula}");
        }
        for formula in [
            "SUM(A1)",
            "\"NOW()\"",
            "NOW",
            "XNOW()",
            "NOWADAYS()",
            "'NOW()'!A1",
            "Sheet1.NOW()",
            "_xlfn.NOW()",
            "RANDARRAY_X()",
            "",
        ] {
            assert!(!recalculated_on_load(formula), "{formula}");
        }
    }

    fn bits(text: &str) -> u64 {
        number(text).to_bits()
    }

    #[test]
    fn numbers_read_the_longest_decimal_prefix() {
        for (text, expected) in [
            ("1_0", 1.0),
            ("1d5", 1.0),
            ("1e", 1.0),
            ("1 2", 1.0),
            ("1,5", 1.0),
            ("0x10", 0.0),
            (".", 0.0),
            ("+", 0.0),
            (".e1", 0.0),
            ("nan", 0.0),
            ("inf", 0.0),
            ("Infinity", 0.0),
            ("1e-999", 0.0),
            ("  ", 0.0),
            ("+5", 5.0),
            ("-.5", -0.5),
            ("5.", 5.0),
            ("1E+2", 100.0),
            ("1e+", 1.0),
            ("0.5.5", 0.5),
            ("1-", 1.0),
            ("1\u{a0}", 1.0),
            ("\u{a0}1", 0.0),
            (" 1 ", 1.0),
            ("1e5x", 100_000.0),
            ("1.#QNAN", 1.0),
            ("1.#IND", 1.0),
            ("123456789012345678901234567890", 1.234_567_890_123_456_8e29),
        ] {
            assert_eq!(bits(text), f64::to_bits(expected), "{text:?}");
        }
    }

    #[test]
    fn signed_zero_survives_a_number_without_digits() {
        assert_eq!(bits("-"), (-0.0_f64).to_bits());
        assert_eq!(bits("--1"), (-0.0_f64).to_bits());
        assert_eq!(bits("-0"), (-0.0_f64).to_bits());
        assert_eq!(bits("+NaN"), 0.0_f64.to_bits());
        assert_eq!(bits("-NaN"), (-0.0_f64).to_bits());
    }

    #[test]
    fn the_words_for_infinity_and_not_a_number_are_case_sensitive() {
        assert!(number("NaN").is_nan());
        assert!(number(" NaN").is_nan());
        assert!(number("NaNx").is_nan());
        assert!(number("NaN1").is_nan());
        assert_eq!(bits("NAN"), 0.0_f64.to_bits());
        assert_eq!(bits("Nan"), 0.0_f64.to_bits());
        for (text, expected) in [
            ("INF", f64::INFINITY),
            ("+INF", f64::INFINITY),
            ("-INF", f64::NEG_INFINITY),
            ("INFx", f64::INFINITY),
            ("INF1", f64::INFINITY),
            ("\tINF", f64::INFINITY),
            ("1.#INF", f64::INFINITY),
            ("-1.#INF", f64::NEG_INFINITY),
            ("1e999", f64::INFINITY),
            ("-1e999", f64::NEG_INFINITY),
            ("1e400", f64::INFINITY),
        ] {
            assert_eq!(number(text).to_bits(), expected.to_bits(), "{text:?}");
        }
        assert_eq!(bits("Inf"), 0.0_f64.to_bits());
    }

    #[test]
    fn external_indices_are_the_bracketed_numbers_outside_quotes() {
        assert_eq!(
            external_indices("SUM([1]S1!A1,[12]S!B2)+'[3]x'"),
            vec![1, 12]
        );
        assert_eq!(external_indices("S1!A1"), Vec::<usize>::new());
    }

    #[test]
    fn only_a_plain_reference_into_another_document_is_external() {
        assert!(is_external_reference("[1]S1!$A$1"));
        assert!(is_external_reference("[2]Sheet 1!A1:$B$2"));
        assert!(!is_external_reference("SUM([1]S1!$A$1,1)"));
        assert!(!is_external_reference("S1!$A$1"));
        assert!(!is_external_reference("[x]S1!A1"));
    }

    #[test]
    fn booleans_are_the_word_true_or_a_non_zero_number() {
        for text in [
            "true", "TRUE", "True", "tRuE", "1", "2", "NaN", "INF", "-INF", "1e999", " 1 ", "1abc",
            "+1", "-1", ".5",
        ] {
            assert!(boolean(text), "{text:?}");
        }
        for text in [
            "false", "FALSE", "0", "0.0", "-0", "0e0", "yes", "T", "t", "inf", "nan", " ", "TRUE ",
            "abc1", "0abc", "0x1", "T R U E", "Infinity",
        ] {
            assert!(!boolean(text), "{text:?}");
        }
    }

    #[test]
    fn shared_string_indexes_read_an_integer_prefix() {
        for (text, expected) in [
            ("0", Some(0)),
            ("7", Some(7)),
            ("abc", Some(0)),
            (" 0 ", Some(0)),
            ("+0", Some(0)),
            ("0x0", Some(0)),
            ("00", Some(0)),
            ("-0", Some(0)),
            ("0 ", Some(0)),
            ("1e0", Some(1)),
            ("1.0", Some(1)),
            ("2147483647", Some(2_147_483_647)),
            ("2147483648", Some(0)),
            ("4294967296", Some(0)),
            ("99999999999", Some(0)),
            ("-1", None),
        ] {
            assert_eq!(shared_string_index(text), expected, "{text:?}");
        }
    }

    #[test]
    fn constant_errors_follow_the_formula_they_would_be() {
        let other = |text: &str| ErrorKind::Other(text.into());
        for (text, expected) in [
            ("#NULL!", ErrorKind::Null),
            ("#div/0!", ErrorKind::DivZero),
            (" #DIV/0! ", ErrorKind::DivZero),
            ("#VALUE!", ErrorKind::Value),
            ("#NUM!", ErrorKind::Num),
            ("#n/a", ErrorKind::Na),
            ("#N/A ", ErrorKind::Na),
            ("#NAME?", ErrorKind::Name),
            ("#REF!", ErrorKind::Name),
            ("#ref!", ErrorKind::Name),
            ("#FOO!", ErrorKind::Name),
            ("#CALC!", ErrorKind::Name),
            ("#NAME", ErrorKind::Name),
            ("Err:502", ErrorKind::Name),
            ("x", ErrorKind::Name),
            ("#SPILL!", other("#SPILL!")),
            ("#Spill!", other("#SPILL!")),
            ("#GETTING_DATA", other("Err:509")),
            ("#getting_data", other("Err:509")),
            ("#NULL!x", other("Err:509")),
        ] {
            assert_eq!(constant_error(text).0, expected, "{text:?}");
        }
    }

    #[test]
    fn references_are_accepted_as_the_oracle_accepts_them() {
        for accepted in [
            "A1:A2",
            "A1",
            "$A$1:$A$2",
            "a1:a2",
            "A1:A2 ",
            "A1:A2  ",
            "A1:B1",
            "1:2",
            "A:A",
            "A1:A2000000",
            "A1:A99999999999999999999",
            "$A1:A$2",
            "A1:a$2",
            "A01:A2",
            "A1:XFD1048576",
            "A2:A1",
            "$A:$A",
            "$1:$2",
            "a:a",
            "A1:A+2",
            "A1:A-1",
            "A1:A+0",
            "A+1",
            "A+1:A2",
            "A1:A0x2",
            "A1:A",
            "A1:$A",
            "A1:B",
            "A1:XFD0",
            "A1:B0z",
            "A1:A$",
            "A1:A0:A2",
            "A1:A1048577x",
            "A1:A-x",
            "A$1:A0",
        ] {
            assert!(
                area(accepted).is_some_and(|area| area.contains(0, 0)),
                "{accepted:?}"
            );
        }
        for rejected in [
            "",
            "A1:",
            ":A2",
            "A1:A2:A3",
            "A1:A2 B1",
            " A1:A2",
            "A1: A2",
            "A1 :A2",
            "A1:A2;",
            "A1;A2",
            "A1:A2,B1",
            "A1:A2x",
            "A1:XFE2",
            "A0:A2",
            "A:XFE",
            "A:A1",
            "1:A2",
            "AAAA1:A2",
            "$$A1:A2",
            "0:2",
            "R1C1:R2C1",
            "Sheet1!A1:A2",
            "A",
            "1",
            "A1:A2\t",
            "A1:A2\u{a0}",
            "A1:A2\n",
            "A-1:A2",
            "A+0:A1",
            "A-0:A1",
            "A1:A+1x",
            "A1:1",
            "A1:0",
            "A1:-1",
            "A1:+1",
            "A1:+A1",
            "A1:!",
            "A1:$",
            "A1:XFE",
            "A1:XFE0",
            "A1:AAAA0",
        ] {
            assert!(area(rejected).is_none(), "{rejected:?}");
        }
        assert!(!area("A2:A3").is_some_and(|area| area.contains(0, 0)));
        assert!(area("A2:A3").is_some_and(|area| area.contains(2, 0)));
    }
}

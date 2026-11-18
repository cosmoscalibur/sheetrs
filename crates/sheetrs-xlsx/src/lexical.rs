//! Reading of cell contents the way `LibreOffice` 26.8 loads them, including its
//! behaviour for text that is not what the cell type announces.

use std::borrow::Cow;

use sheetrs_core::{DateSystem, ErrorKind};

/// The whitespace skipped before a stored number.
const NUMBER_WHITESPACE: [char; 2] = [' ', '\t'];
/// The whitespace of XML 1.0.
const XML_WHITESPACE: [char; 4] = [' ', '\t', '\r', '\n'];

/// The whitespace skipped around a shared string reference and a date: XML whitespace, the
/// spaces U+2000 to U+200B and the line and paragraph separators.
fn is_wide_whitespace(character: char) -> bool {
    matches!(
        character,
        ' ' | '\t' | '\r' | '\n' | '\u{2000}'..='\u{200B}' | '\u{2028}' | '\u{2029}'
    )
}

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

/// The longest prefix of `text` of the form `digits[.digits][e[+-]digits]` or `.digits[...]`,
/// if it has at least one digit in its mantissa.
fn decimal_prefix(text: &str) -> Option<&str> {
    let bytes = text.as_bytes();
    let count_digits = |from: usize| {
        bytes.get(from..).map_or(0, |tail| {
            tail.iter().take_while(|byte| byte.is_ascii_digit()).count()
        })
    };
    let integer = count_digits(0);
    let mut end = integer;
    let mut mantissa_digits = integer;
    if bytes.get(end) == Some(&b'.') {
        let fraction = count_digits(end + 1);
        end += 1 + fraction;
        mantissa_digits += fraction;
    }
    if mantissa_digits == 0 {
        return None;
    }
    if matches!(bytes.get(end), Some(b'e' | b'E')) {
        let sign = usize::from(matches!(bytes.get(end + 1), Some(b'+' | b'-')));
        let exponent = count_digits(end + 1 + sign);
        if exponent > 0 {
            end += 1 + sign + exponent;
        }
    }
    text.get(..end)
}

/// A stored boolean: `true` in any case, or any number other than zero.
pub(crate) fn boolean(text: &str) -> bool {
    if text.eq_ignore_ascii_case("true") {
        return true;
    }
    let value = number(text);
    value.is_nan() || value.abs() > 0.0
}

/// Whether the `date1904` attribute of the workbook properties is on: the words `true`,
/// `on` and `t`, or a number other than zero.
pub(crate) fn date_1904_flag(text: &str) -> bool {
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

/// The integer a stored text starts with, after the whitespace `skip` accepts and an optional
/// sign; text without digits, or a number beyond the 32-bit range, is zero.
fn integer_prefix_skipping(text: &str, skip: fn(char) -> bool) -> i64 {
    let rest = text.trim_start_matches(skip);
    let (negative, unsigned) = match rest.strip_prefix('-') {
        Some(after) => (true, after),
        None => (false, rest.strip_prefix('+').unwrap_or(rest)),
    };
    let digits_end = unsigned.bytes().take_while(u8::is_ascii_digit).count();
    let magnitude = unsigned
        .get(..digits_end)
        .and_then(|digits| digits.parse::<i64>().ok())
        .unwrap_or(0);
    let value = if negative { -magnitude } else { magnitude };
    if i32::try_from(value).is_ok() {
        value
    } else {
        0
    }
}

fn is_xml_whitespace(character: char) -> bool {
    XML_WHITESPACE.contains(&character)
}

/// The integer of an attribute such as a row number or a shared formula group, which skips
/// only XML whitespace.
pub(crate) fn integer_prefix(text: &str) -> i64 {
    integer_prefix_skipping(text, is_xml_whitespace)
}

/// The index a stored shared-string reference stands for, or `None` when it is negative.
pub(crate) fn shared_string_index(text: &str) -> Option<usize> {
    usize::try_from(integer_prefix_skipping(text, is_wide_whitespace)).ok()
}

/// The error a cell of type `e` without a formula holds.
///
/// The six classic errors are matched ignoring case and surrounding whitespace; other text
/// is read as the formula it would be, so most of it is `#NAME?`.
pub(crate) fn constant_error(text: &str) -> ErrorKind {
    const CLASSIC: [(&str, ErrorKind); 6] = [
        ("#NULL!", ErrorKind::Null),
        ("#DIV/0!", ErrorKind::DivZero),
        ("#VALUE!", ErrorKind::Value),
        ("#NUM!", ErrorKind::Num),
        ("#N/A", ErrorKind::Na),
        ("#NAME?", ErrorKind::Name),
    ];
    let upper = text.trim_matches(XML_WHITESPACE).to_ascii_uppercase();
    for (name, kind) in CLASSIC {
        if upper == name {
            return kind;
        }
        if upper.starts_with(name) {
            return ErrorKind::Other("Err:509".into());
        }
    }
    match upper.as_str() {
        "#SPILL!" => ErrorKind::Other("#SPILL!".into()),
        "#GETTING_DATA" => ErrorKind::Other("Err:509".into()),
        _ => ErrorKind::Name,
    }
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

/// The serial number of an ISO 8601 date or date-time (`2020-01-02`, `2020-01-02T10:30:00+01:00`),
/// or `None` when the text is not one.
///
/// The year has one or more digits, an optional minus sign and a value from -32768 to 32767
/// other than 0. Fractional seconds are dropped. `Z` is UTC; a positive offset counts its hours
/// only and a negative one counts its minutes twice, up to 14 hours.
pub(crate) fn iso_date(text: &str, system: DateSystem) -> Option<f64> {
    let text = text.trim_matches(is_wide_whitespace);
    let (year, rest) = split_year(text)?;
    let (month, rest) = fixed_digits(rest.strip_prefix('-')?, 2)?;
    let (day, rest) = fixed_digits(rest.strip_prefix('-')?, 2)?;
    let days = system.serial_of_date(year, u32::try_from(month).ok()?, u32::try_from(day).ok()?)?;

    let (seconds_of_day, rest) = match rest.strip_prefix(['T', 't']) {
        Some(time) => time_of_day(time)?,
        None => (0, rest),
    };
    let offset = utc_offset(rest)?;
    let seconds = seconds_of_day - offset;
    let day_shift = seconds.div_euclid(86_400);
    let time = seconds.rem_euclid(86_400);
    // Summed in this order, which is the one that reproduces `LibreOffice` bit for bit.
    let whole_days = f64::from(i32::try_from(days + day_shift).ok()?);
    let fraction = |count: i64, per_day: f64| Some(f64::from(i32::try_from(count).ok()?) / per_day);
    Some(
        whole_days
            + fraction(time / 3600, 24.0)?
            + fraction(time % 3600 / 60, 1440.0)?
            + fraction(time % 60, 86_400.0)?,
    )
}

fn split_year(text: &str) -> Option<(i64, &str)> {
    let (negative, unsigned) = match text.strip_prefix('-') {
        Some(after) => (true, after),
        None => (false, text),
    };
    let digits = unsigned.bytes().take_while(u8::is_ascii_digit).count();
    let (year_digits, rest) = unsigned.split_at(digits);
    let significant = year_digits.trim_start_matches('0');
    if year_digits.is_empty() || significant.len() > 5 {
        return None;
    }
    let magnitude: i64 = if significant.is_empty() {
        0
    } else {
        significant.parse().ok()?
    };
    let year = if negative { -magnitude } else { magnitude };
    ((-32_768..=32_767).contains(&year) && year != 0).then_some((year, rest))
}

/// `count` ASCII digits as a number, and what follows them.
fn fixed_digits(text: &str, count: usize) -> Option<(i64, &str)> {
    let digits = text.get(..count)?;
    if !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    Some((digits.parse().ok()?, text.get(count..)?))
}

/// `hh:mm:ss` with optional fractional seconds, as seconds of the day, and what follows.
///
/// The fraction needs a `.` or `,` and at least one digit. Hour 24 is midnight at the end of
/// the day and takes no fraction in its first nine digits.
fn time_of_day(text: &str) -> Option<(i64, &str)> {
    let (hours, rest) = fixed_digits(text, 2)?;
    let (minutes, rest) = fixed_digits(rest.strip_prefix(':')?, 2)?;
    let (seconds, rest) = fixed_digits(rest.strip_prefix(':')?, 2)?;
    if hours > 24 || minutes > 59 || seconds > 59 {
        return None;
    }
    let (fraction_digits, rest) = match rest.strip_prefix(['.', ',']) {
        Some(fraction) => {
            let digits = fraction.bytes().take_while(u8::is_ascii_digit).count();
            if digits == 0 {
                return None;
            }
            fraction.split_at(digits)
        }
        None => ("", rest),
    };
    let leading_fraction_is_zero = fraction_digits.bytes().take(9).all(|digit| digit == b'0');
    if hours == 24 && (minutes + seconds > 0 || !leading_fraction_is_zero) {
        return None;
    }
    Some((hours * 3600 + minutes * 60 + seconds, rest))
}

/// The offset in seconds of a trailing `Z`, `+hh:mm`, `-hh:mm` or nothing, as the oracle counts it.
fn utc_offset(text: &str) -> Option<i64> {
    if text.is_empty() || text == "Z" || text == "z" {
        return Some(0);
    }
    let positive = match text.get(..1)? {
        "+" => true,
        "-" => false,
        _ => return None,
    };
    let (hours, rest) = fixed_digits(text.get(1..)?, 2)?;
    let (minutes, rest) = fixed_digits(rest.strip_prefix(':')?, 2)?;
    if !rest.is_empty() || minutes > 59 || hours > 14 || (hours == 14 && minutes > 0) {
        return None;
    }
    Some(if positive {
        hours * 3600
    } else {
        -(hours * 3600 + 2 * minutes * 60)
    })
}

#[cfg(test)]
mod tests {
    use super::*;

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
            assert_eq!(constant_error(text), expected, "{text:?}");
        }
    }

    #[test]
    fn dates_become_serial_numbers() {
        let serial = |text| iso_date(text, DateSystem::Date1900).map(f64::to_bits);
        for (text, expected) in [
            ("2020-01-02", "43832.0"),
            ("2020-01-02T10:30:00", "43832.4375"),
            ("2020-01-02T10:30:00.5", "43832.4375"),
            ("2020-01-02T10:30:00Z", "43832.4375"),
            ("2020-01-02T10:30:00.999", "43832.4375"),
            ("2020-01-02T23:59:59.999", "43832.99998842593"),
            ("2020-01-02T10:30:00+01:00", "43832.395833333336"),
            ("2020-01-02T10:30:00-05:00", "43832.645833333336"),
            ("2020-01-02T00:00:00+01:00", "43831.958333333336"),
            ("2020-01-02+01:00", "43831.958333333336"),
            ("2020-01-02Z", "43832.0"),
            ("2020-01-02z", "43832.0"),
            ("2020-01-02t10:30:00", "43832.4375"),
            ("  2020-01-02", "43832.0"),
            ("2020-01-02 ", "43832.0"),
            ("12020-01-02", "3696257.0"),
            ("2020-01-02T24:00:00", "43833.0"),
            ("1900-01-01", "2.0"),
            ("1900-03-01", "61.0"),
            ("1899-12-31", "1.0"),
            ("0001-01-01", "-693593.0"),
            ("9999-12-31", "2958465.0"),
        ] {
            let expected: f64 = expected.parse().unwrap_or(f64::NAN);
            assert_eq!(serial(text), Some(expected.to_bits()), "{text:?}");
        }
    }

    #[test]
    fn text_that_is_not_an_iso_date_has_no_serial() {
        for text in [
            "10:30:00",
            "2020-1-2",
            "20200102",
            "abc",
            "2020-02-30",
            "2020-01-02T10:30",
            "+2020-01-02",
            "2020-01-02T10:60:00",
            "2020-01-02T10:30:60",
            "2019-02-29",
            "1900-02-29",
            "2020-13-01",
            "2020-00-10",
            " ",
            "",
        ] {
            assert_eq!(iso_date(text, DateSystem::Date1900), None, "{text:?}");
        }
    }

    #[test]
    fn the_1904_system_counts_from_1904() {
        let serial = |text| iso_date(text, DateSystem::Date1904);
        assert_eq!(serial("1904-01-01"), Some(0.0));
        assert_eq!(serial("2020-01-02"), Some(42_370.0));
        assert_eq!(serial("1900-01-01"), Some(-1_460.0));
    }
}

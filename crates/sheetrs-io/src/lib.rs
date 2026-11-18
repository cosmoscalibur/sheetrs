//! Lexical rules that the format readers share, as `LibreOffice` 26.8 applies them to the
//! text of a file.

use sheetrs_core::DateSystem;

/// The whitespace skipped around a shared string reference and a date: XML whitespace, the
/// spaces U+2000 to U+200B and the line and paragraph separators.
#[must_use]
pub fn is_wide_whitespace(character: char) -> bool {
    matches!(
        character,
        ' ' | '\t' | '\r' | '\n' | '\u{2000}'..='\u{200B}' | '\u{2028}' | '\u{2029}'
    )
}

/// The longest prefix of `text` of the form `digits[.digits][e[+-]digits]` or `.digits[...]`,
/// if it has at least one digit in its mantissa.
#[must_use]
pub fn decimal_prefix(text: &str) -> Option<&str> {
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

/// What a date-time does with the fraction of its seconds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Fraction {
    /// The fraction is dropped.
    Dropped,
    /// The fraction counts in the serial number.
    Kept,
}

/// How a format reads the text of a date-time.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DateRules {
    /// What happens to the fraction of the seconds.
    pub fraction: Fraction,
    /// Whether a year before 1 counts as in the calendar without a year 0, where -1 is the year
    /// before 1, or as the astronomical year of that number.
    pub no_year_zero: bool,
}

impl DateRules {
    /// The dates of an XLSX cell of type `d`.
    pub const XLSX: Self = Self {
        fraction: Fraction::Dropped,
        no_year_zero: false,
    };
    /// The dates of an ODS `office:date-value`.
    pub const ODS: Self = Self {
        fraction: Fraction::Kept,
        no_year_zero: true,
    };
}

/// The day and the time of day of an ISO 8601 date-time, as `iso_date` reads them.
struct DateTime<'a> {
    /// The serial number of the day, after the offset.
    day: i64,
    /// Seconds of the day, after the offset.
    seconds: i64,
    /// The digits after the decimal point of the seconds.
    fraction_digits: &'a str,
}

/// Splits `text` into its day and time of day, or `None` when it is not an ISO 8601 date-time.
fn date_time(text: &str, system: DateSystem, rules: DateRules) -> Option<DateTime<'_>> {
    let text = text.trim_matches(is_wide_whitespace);
    let (year, rest) = split_year(text)?;
    let (month, rest) = fixed_digits(rest.strip_prefix('-')?, 2)?;
    let (day, rest) = fixed_digits(rest.strip_prefix('-')?, 2)?;
    let astronomical_year = if rules.no_year_zero && year < 0 {
        year + 1
    } else {
        year
    };
    let days = system.serial_of_date(
        astronomical_year,
        u32::try_from(month).ok()?,
        u32::try_from(day).ok()?,
    )?;
    let (seconds_of_day, fraction_digits, rest) = match rest.strip_prefix(['T', 't']) {
        Some(time) => time_of_day(time)?,
        None => (0, "", rest),
    };
    let seconds = seconds_of_day - utc_offset(rest)?;
    Some(DateTime {
        day: days + seconds.div_euclid(86_400),
        seconds: seconds.rem_euclid(86_400),
        fraction_digits,
    })
}

/// The serial number of an ISO 8601 date or date-time (`2020-01-02`, `2020-01-02T10:30:00+01:00`),
/// or `None` when the text is not one.
///
/// The year has one or more digits, an optional minus sign and a value from -32768 to 32767
/// other than 0. Fractional seconds are dropped or kept as the rules say. `Z` is UTC; a
/// positive offset counts its hours only and a negative one counts its minutes twice, up to 14
/// hours.
#[must_use]
pub fn iso_date(text: &str, system: DateSystem, rules: DateRules) -> Option<f64> {
    let DateTime {
        day,
        seconds: time,
        fraction_digits,
    } = date_time(text, system, rules)?;
    // Summed in this order, which is the one that reproduces `LibreOffice` bit for bit.
    let whole_days = f64::from(i32::try_from(day).ok()?);
    let fraction_of =
        |count: i64, per_day: f64| Some(f64::from(i32::try_from(count).ok()?) / per_day);
    let seconds = match rules.fraction {
        Fraction::Dropped => f64::from(i32::try_from(time % 60).ok()?),
        Fraction::Kept => {
            let digits = fraction_digits
                .get(..fraction_digits.len().min(9))
                .unwrap_or_default();
            let tail: f64 = format!("0.{digits}").parse().unwrap_or(0.0);
            f64::from(i32::try_from(time % 60).ok()?) + if digits.is_empty() { 0.0 } else { tail }
        }
    };
    Some(
        whole_days
            + fraction_of(time / 3600, 24.0)?
            + fraction_of(time % 3600 / 60, 1440.0)?
            + seconds / 86_400.0,
    )
}

/// The serial number of the day of an ISO 8601 date or date-time, without its time of day.
#[must_use]
pub fn iso_day(text: &str, system: DateSystem, rules: DateRules) -> Option<i64> {
    date_time(text, system, rules).map(|date_time| date_time.day)
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
/// The fraction needs a `.` or `,` and at least one digit, and is returned as its digits. Hour
/// 24 is midnight at the end of the day and takes no fraction in its first nine digits.
fn time_of_day(text: &str) -> Option<(i64, &str, &str)> {
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
    Some((hours * 3600 + minutes * 60 + seconds, fraction_digits, rest))
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

/// The whitespace of XML 1.0.
pub const XML_WHITESPACE: [char; 4] = [' ', '\t', '\r', '\n'];

/// The integer a stored text starts with, after the whitespace `skip` accepts and an optional
/// sign; text without digits, or a number beyond the 32-bit range, is zero.
#[must_use]
pub fn integer_prefix_skipping(text: &str, skip: fn(char) -> bool) -> i64 {
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
#[must_use]
pub fn integer_prefix(text: &str) -> i64 {
    integer_prefix_skipping(text, is_xml_whitespace)
}

/// Whether `formula` calls one of `functions`: a name, in any case, followed by an opening
/// parenthesis that spaces may separate from it. Names inside text and inside quoted sheet names
/// are not calls.
#[must_use]
pub fn calls_any(formula: &str, functions: &[&str]) -> bool {
    let bytes = formula.as_bytes();
    let is_name = |byte: &u8| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.');
    let mut at = 0;
    while let Some(&byte) = bytes.get(at) {
        match byte {
            b'"' | b'\'' => {
                // Text, or a quoted name: it ends at the first quote that is not doubled.
                at += 1;
                while let Some(&inner) = bytes.get(at) {
                    at += 1;
                    if inner == byte {
                        if bytes.get(at) == Some(&byte) {
                            at += 1;
                        } else {
                            break;
                        }
                    }
                }
            }
            _ if is_name(&byte) => {
                let length = bytes
                    .get(at..)
                    .map_or(0, |tail| tail.iter().take_while(|b| is_name(b)).count());
                let name = bytes.get(at..at + length).unwrap_or_default();
                let after = bytes.get(at + length..).map_or(&[][..], |tail| {
                    let spaces = tail.iter().take_while(|b| **b == b' ').count();
                    tail.get(spaces..).unwrap_or_default()
                });
                if after.first() == Some(&b'(')
                    && functions
                        .iter()
                        .any(|function| function.as_bytes().eq_ignore_ascii_case(name))
                {
                    return true;
                }
                at += length;
            }
            _ => at += 1,
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn dates_become_serial_numbers() {
        let serial = |text| iso_date(text, DateSystem::Date1900, DateRules::XLSX).map(f64::to_bits);
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
            assert_eq!(
                iso_date(text, DateSystem::Date1900, DateRules::XLSX),
                None,
                "{text:?}"
            );
        }
    }

    #[test]
    fn ods_dates_keep_fractions_and_count_years_before_one_without_a_zero() {
        let serial = |text| iso_date(text, DateSystem::Date1900, DateRules::ODS);
        assert_eq!(
            serial("2020-01-02T12:00:00.5"),
            Some(43_832.500_005_787_035)
        );
        assert_eq!(serial("0001-01-01"), Some(-693_593.0));
        assert_eq!(serial("-0001-01-01"), Some(-693_959.0));
        assert_eq!(serial("-0002-01-01"), Some(-694_324.0));
        let literal = |text| iso_date(text, DateSystem::Date1900, DateRules::XLSX);
        assert_eq!(literal("-0001-01-01"), Some(-694_324.0));
    }

    #[test]
    fn the_1904_system_counts_from_1904() {
        let serial = |text| iso_date(text, DateSystem::Date1904, DateRules::XLSX);
        assert_eq!(serial("1904-01-01"), Some(0.0));
        assert_eq!(serial("2020-01-02"), Some(42_370.0));
        assert_eq!(serial("1900-01-01"), Some(-1_460.0));
    }
}

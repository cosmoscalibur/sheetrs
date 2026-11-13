//! The dates of a workbook as dates: the `datetime` option of a query.
//!
//! A number is a date, a time or both when its format says so (the recognition [`display`] uses:
//! `yyyy-mm-dd`, `hh:mm:ss`, `dd/mm/yyyy hh:mm`, `d-mmm`, `mm:ss`…). What it becomes comes from the
//! serial number and the date system of the file, in the calendar `LibreOffice` uses (the Gregorian one
//! from 1582-10-15 and the Julian one before it; the serial 0 of the usual system is 1899-12-30, which is
//! how `LibreOffice` reads an XLSX too: it has no 1900 leap-year day):
//!
//! - a date format: `YYYY-MM-DD` from the whole part of the serial;
//! - a date and time format: `YYYY-MM-DDThh:mm:ss` from the serial with the seconds cut, as `LibreOffice`
//!   shows them (`23:59:59` for `0.999999` of a day); when the format writes fractions of a second
//!   (`ss.000`) they are rounded to that many digits, up to three, and the text has them, never carrying into
//!   the next second;
//! - a time format: `hh:mm:ss` from the fractional part, the day dropped (the way a time format shows a
//!   serial greater than one);
//! - an elapsed-time format (`[h]:mm:ss`, which counts hours past 24) is a duration and not a time of
//!   day: it stays a number;
//! - a serial whose year is outside 1 to 9999 stays a number: ISO 8601 writes four digits and a
//!   `datetime.date` stops there, where `LibreOffice` goes on to the year 32767.
//!
//! [`display`]: crate::text::TextContext::display

use crate::text::{Kind, kind};
use sheetrs_engine::dates::calendar_from_days;
use sheetrs_engine::options::DateSystem;
use sheetrs_protocol::DateValue;
use std::fmt::Write as _;

/// How many digits of a second the format writes after the point (`ss.000` is 3), at most 3.
fn fraction_digits(format: Option<&str>) -> u32 {
    let Some(format) = format else {
        return 0;
    };
    let mut chars = format.chars().peekable();
    let mut after_seconds = false;
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                for q in chars.by_ref() {
                    if q == '"' {
                        break;
                    }
                }
            }
            '\\' | '_' | '*' => {
                chars.next();
            }
            's' | 'S' => after_seconds = true,
            '.' if after_seconds => {
                let mut zeros = 0;
                while chars.peek() == Some(&'0') {
                    chars.next();
                    zeros += 1;
                }
                return zeros.min(3);
            }
            _ => after_seconds = false,
        }
    }
    0
}

/// Seconds of tolerance for the error of a serial: a time typed as `00:15:00` is stored as a number whose
/// seconds come out at `899.9999999` and `LibreOffice` shows `00:15:00`.
const EPSILON: f64 = 1e-6;

/// The whole days of `x` and the ticks of the day (`10^digits` per second). `LibreOffice` cuts the whole
/// seconds (`23:59:59` for `0.999999` of a day, never `24:00:00`) and rounds the fraction it writes
/// (`.914` for `913.6` milliseconds) without letting it reach the next second (`.99` for `99.9`
/// hundredths): neither the second nor the day is ever carried.
fn split(x: f64, digits: u32) -> (f64, i64) {
    let scale = 10_i64.pow(digits);
    let days = x.floor();
    let seconds = (x - days) * 86_400.0;
    let whole = (seconds + EPSILON).floor().min(86_399.0);
    let fraction = ((seconds - whole).max(0.0) * scale as f64).round() as i64;
    (days, whole as i64 * scale + fraction.min(scale - 1))
}

fn clock(ticks: i64, digits: u32) -> String {
    let scale = 10_i64.pow(digits);
    let seconds = ticks / scale;
    let mut text = format!(
        "{:02}:{:02}:{:02}",
        seconds / 3600,
        seconds / 60 % 60,
        seconds % 60
    );
    if digits > 0 {
        let _ = write!(text, ".{:0width$}", ticks % scale, width = digits as usize);
    }
    text
}

/// The calendar date of a whole serial, as `LibreOffice` writes it: the Gregorian calendar from
/// 1582-10-15 and the Julian one before it (the calendar of ICU, which `LibreOffice` uses). The years 1 to
/// 9999, the ones ISO 8601 writes with four digits (and a `datetime.date` holds), are the dates;
/// `LibreOffice` shows more, up to the year 32767.
fn ymd(day: f64, sys: DateSystem) -> Option<String> {
    if !day.is_finite() || day.abs() > 1e8 {
        return None;
    }
    let days = sys.days_from_unix_epoch() + day as i64;
    let (y, m, d) = calendar_from_days(days);
    if !(1..=9999).contains(&y) {
        return None;
    }
    Some(format!("{y:04}-{m:02}-{d:02}"))
}

/// The number `x` with the format `format` as a date, a time or both; `None` when it is none of them
/// (the cell keeps its number).
#[must_use]
pub(crate) fn typed(x: f64, format: Option<&str>, sys: DateSystem) -> Option<DateValue> {
    if !x.is_finite() {
        return None;
    }
    match kind(format) {
        Kind::Date { time: false, .. } => ymd(x.floor(), sys).map(DateValue::Date),
        Kind::Date { time: true, .. } => {
            let digits = fraction_digits(format);
            let (days, ticks) = split(x, digits);
            let date = ymd(days, sys)?;
            Some(DateValue::Datetime(format!(
                "{date}T{}",
                clock(ticks, digits)
            )))
        }
        Kind::Time { total_hours: false } => {
            let digits = fraction_digits(format);
            let (_, ticks) = split(x, digits);
            Some(DateValue::Time(clock(ticks, digits)))
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const S1900: DateSystem = DateSystem::Null1899;

    fn t(x: f64, f: &str) -> Option<String> {
        typed(x, Some(f), S1900).map(|v| v.to_json().to_string())
    }

    #[test]
    fn a_date_a_time_and_both() {
        assert_eq!(
            t(45_366.0, "yyyy-mm-dd").unwrap(),
            r#"{"type":"date","value":"2024-03-15"}"#
        );
        assert_eq!(
            t(45_366.437_5, "yyyy-mm-dd hh:mm:ss").unwrap(),
            r#"{"type":"datetime","value":"2024-03-15T10:30:00"}"#
        );
        assert_eq!(
            t(45_366.437_5, "hh:mm").unwrap(),
            r#"{"type":"time","value":"10:30:00"}"#
        );
        // The day of a date format is the whole part: the time is not rounded up into it.
        assert_eq!(
            t(45_366.999_999, "yyyy-mm-dd").unwrap(),
            r#"{"type":"date","value":"2024-03-15"}"#
        );
    }

    #[test]
    fn seconds_are_cut_and_fractions_rounded_without_a_carry() {
        assert_eq!(
            t(45_366.999_999, "yyyy-mm-dd hh:mm:ss").unwrap(),
            r#"{"type":"datetime","value":"2024-03-15T23:59:59"}"#
        );
        // The fraction is rounded but never carried into the next second.
        assert_eq!(
            t(45_366.999_999_999_9, "yyyy-mm-dd hh:mm:ss.000").unwrap(),
            r#"{"type":"datetime","value":"2024-03-15T23:59:59.999"}"#
        );
        // 00:15:00 typed with ten digits.
        assert_eq!(
            t(45_366.010_416_666_7, "hh:mm:ss").unwrap(),
            r#"{"type":"time","value":"00:15:00"}"#
        );
        assert_eq!(
            t(0.5, "hh:mm:ss").unwrap(),
            r#"{"type":"time","value":"12:00:00"}"#
        );
    }

    #[test]
    fn fractions_of_a_second_are_kept_when_the_format_writes_them() {
        assert_eq!(
            t(45_366.0 + 3723.456 / 86_400.0, "yyyy-mm-dd hh:mm:ss.000").unwrap(),
            r#"{"type":"datetime","value":"2024-03-15T01:02:03.456"}"#
        );
        assert_eq!(fraction_digits(Some("hh:mm:ss.00")), 2);
        assert_eq!(fraction_digits(Some("hh:mm:ss")), 0);
        assert_eq!(fraction_digits(Some("mm.00")), 0);
    }

    #[test]
    fn what_is_not_a_date_stays_a_number() {
        assert_eq!(t(45_366.0, "General"), None);
        assert_eq!(t(0.25, "0.00%"), None);
        assert_eq!(t(1.5, "[h]:mm:ss"), None);
        assert_eq!(t(1e12, "yyyy-mm-dd"), None);
        assert_eq!(t(2_958_466.0, "yyyy-mm-dd"), None); // the year 10000
        assert_eq!(t(f64::NAN, "yyyy-mm-dd"), None);
        assert_eq!(typed(45_366.0, None, S1900), None);
    }

    #[test]
    fn the_date_system_of_the_file_numbers_the_dates() {
        let s1904 = DateSystem::Null1904;
        assert_eq!(
            typed(43904.0, Some("yyyy-mm-dd"), s1904)
                .unwrap()
                .text()
                .to_string(),
            "2024-03-15"
        );
    }
}

//! The ISO 8601 dates and date-times that the text filter reads as numbers even when it does
//! not detect special numbers.

use sheetrs_core::DateSystem;

/// Days from 1899-12-30 back to the Julian day number 0, as an offset to subtract.
const JULIAN_DAY_OF_SERIAL_ZERO: i64 = 2_415_019;

/// The serial number of `text`, a date `YYYY-MM-DD` with an optional time `Thh:mm:ss`, between
/// spaces; `None` when it is anything else.
///
/// The year has four or five digits and an optional minus sign. A date before the switch to the
/// Gregorian calendar, which skipped 5 to 14 October 1582, counts as a date of the Julian
/// calendar, so the serial numbers run on across the switch; a year before 1 counts as in the
/// calendar without a year 0. Whether February has 29 days follows the Gregorian rule in either
/// calendar.
pub(crate) fn iso_date(text: &str) -> Option<f64> {
    let text = text.trim_matches(' ');
    let (year, rest) = year(text)?;
    let (month, rest) = digits(rest.strip_prefix('-')?, 2)?;
    let (day, rest) = digits(rest.strip_prefix('-')?, 2)?;
    let (seconds, fraction, rest) = match rest.strip_prefix(['T', 't']) {
        Some(time) => time_of_day(time)?,
        None => (0, 0.0, rest),
    };
    if !rest.is_empty() {
        return None;
    }
    let serial = serial(year, u32::try_from(month).ok()?, u32::try_from(day).ok()?)?;
    let time_of_day = (f64::from(i32::try_from(seconds).ok()?) + fraction) / 86_400.0;
    Some(f64::from(i32::try_from(serial).ok()?) + time_of_day)
}

fn year(text: &str) -> Option<(i64, &str)> {
    let (negative, unsigned) = match text.strip_prefix('-') {
        Some(after) => (true, after),
        None => (false, text),
    };
    let count = unsigned.bytes().take_while(u8::is_ascii_digit).count();
    let (digits, rest) = unsigned.split_at(count);
    if !(4..=5).contains(&digits.len()) {
        return None;
    }
    let magnitude: i64 = digits.parse().ok()?;
    let year = if negative { -magnitude } else { magnitude };
    ((-32_767..=32_767).contains(&year) && year != 0).then_some((year, rest))
}

fn digits(text: &str, count: usize) -> Option<(i64, &str)> {
    let head = text.get(..count)?;
    if !head.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    Some((head.parse().ok()?, text.get(count..)?))
}

/// `hh:mm:ss` with optional fractional seconds: the whole seconds of the day, the fraction and
/// what follows. Hour 24 stands for the end of the day and takes no minutes, seconds or
/// fraction.
fn time_of_day(text: &str) -> Option<(i64, f64, &str)> {
    let (hours, rest) = digits(text, 2)?;
    let (minutes, rest) = digits(rest.strip_prefix(':')?, 2)?;
    let (seconds, rest) = digits(rest.strip_prefix(':')?, 2)?;
    if hours > 24 || minutes > 59 || seconds > 59 {
        return None;
    }
    let (fraction, rest) = match rest.strip_prefix(['.', ',']) {
        Some(after) => {
            let count = after.bytes().take_while(u8::is_ascii_digit).count();
            if count == 0 {
                return None;
            }
            let (digits, tail) = after.split_at(count);
            let kept = digits.get(..digits.len().min(9)).unwrap_or_default();
            (format!("0.{kept}").parse::<f64>().ok()?, tail)
        }
        None => (0.0, rest),
    };
    if hours == 24 && (minutes + seconds > 0 || fraction > 0.0) {
        return None;
    }
    Some((hours * 3600 + minutes * 60 + seconds, fraction, rest))
}

fn is_gregorian_leap(year: i64) -> bool {
    year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)
}

/// The serial number of a calendar date, or `None` when it does not exist.
fn serial(year: i64, month: u32, day: u32) -> Option<i64> {
    let astronomical = if year < 0 { year + 1 } else { year };
    let longest = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if is_gregorian_leap(astronomical) => 29,
        2 => 28,
        _ => return None,
    };
    if day == 0 || day > longest {
        return None;
    }
    let switch = (1582, 10, 15);
    let julian = (year, month, day) < switch;
    if julian && (year, month) == (1582, 10) && day > 4 {
        return None;
    }
    if julian {
        let shifted = i64::from(month <= 2);
        let years = astronomical + 4800 - shifted;
        let months = i64::from(month) + 12 * shifted - 3;
        let day_number =
            i64::from(day) + (153 * months + 2) / 5 + 365 * years + years.div_euclid(4) - 32_083;
        Some(day_number - JULIAN_DAY_OF_SERIAL_ZERO)
    } else {
        DateSystem::Date1900.serial_of_date(astronomical, month, day)
    }
}

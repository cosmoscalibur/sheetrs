//! Serial-number date arithmetic for both spreadsheet date systems.

use crate::options::DateSystem;

fn out_of_range<T>() -> Option<T> {
    crate::notes::note(crate::notes::DATE_RANGE);
    None
}

pub use sheetrs_formula::calendar::{
    civil_from_days, days_from_civil, serial_range, serial_to_ymd_quiet,
};

/// The date a day number (days since 1970-01-01) is written as by the number formatter of
/// `LibreOffice` (ICU's calendar): the Gregorian one from 1582-10-15 and the Julian one before it, where
/// the date functions (`YEAR`, `MONTH`, `DAY`) go on with the Gregorian one. The years go on to
/// the negative ones.
#[must_use]
pub fn calendar_from_days(days: i64) -> (i64, i64, i64) {
    if days >= days_from_civil(1582, 10, 15) {
        return civil_from_days(days);
    }
    // The Julian calendar, by the algorithm of Richards; the Julian day number of 1970-01-01 is 2440588.
    let f = days + 2_440_588 + 1401;
    let e = 4 * f + 3;
    let g = e.rem_euclid(1461) / 4;
    let h = 5 * g + 2;
    let day = h.rem_euclid(153) / 5 + 1;
    let month = (h / 153 + 2).rem_euclid(12) + 1;
    let year = e.div_euclid(1461) - 4716 + (14 - month) / 12;
    (year, month, day)
}

/// The first and the last day number (since 1970-01-01) that the number formatter of `LibreOffice` writes
/// as a date: the Julian date -32768-09-06 and 32768-01-01 (the serials -12661859 and 11274307 of the usual
/// system). Outside them it writes `#FMT`.
pub const LIBRE_FIRST_DAY: i64 = -12_661_859 - 25_569;
pub const LIBRE_LAST_DAY: i64 = 11_274_307 - 25_569;

fn epoch(sys: DateSystem) -> i64 {
    sys.days_from_unix_epoch()
}

/// Serial number of a calendar date; months and days may overflow like in
/// `DATE()`. `None` when the result is outside the representable range.
pub fn ymd_to_serial(y: i64, m: i64, d: i64, sys: DateSystem) -> Option<f64> {
    ymd_to_serial_quiet(y, m, d, sys).or_else(out_of_range)
}

/// Like [`ymd_to_serial`], for text that merely may not be a date.
#[must_use]
pub fn ymd_to_serial_quiet(y: i64, m: i64, d: i64, sys: DateSystem) -> Option<f64> {
    let y = y + (m - 1).div_euclid(12);
    let m = (m - 1).rem_euclid(12) + 1;
    let days = days_from_civil(y, m, 1) + (d - 1);
    let serial = days - epoch(sys);
    let (min, max) = serial_range(sys);
    (min..max).contains(&serial).then_some(serial as f64)
}

pub fn serial_to_ymd(serial: f64, sys: DateSystem) -> Option<(i64, i64, i64)> {
    serial_to_ymd_quiet(serial, sys).or_else(out_of_range)
}

/// 0 = Sunday … 6 = Saturday.
#[must_use]
pub fn weekday_sun0(serial: f64, sys: DateSystem) -> i64 {
    let s = serial.floor() as i64;
    // 1970-01-01 was a Thursday (4).
    (epoch(sys) + s + 4).rem_euclid(7)
}

#[must_use]
pub fn unix_to_serial(secs: f64, sys: DateSystem) -> f64 {
    let days = secs / 86_400.0;
    days + (days_from_civil(1970, 1, 1) - epoch(sys)) as f64
}

#[must_use]
pub fn is_leap(y: i64) -> bool {
    (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
}

#[must_use]
pub fn days_in_month(y: i64, m: i64) -> i64 {
    match m {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        _ => {
            if is_leap(y) {
                29
            } else {
                28
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    const NULL_1899: DateSystem = DateSystem::Null1899;
    const NULL_1904: DateSystem = DateSystem::Null1904;

    #[test]
    fn known_serials_1899() {
        // Serial 0 is 1899-12-30 on the real calendar: no 1900-02-29, negative serials.
        assert_eq!(ymd_to_serial(1900, 1, 1, NULL_1899), Some(2.0));
        assert_eq!(ymd_to_serial(1900, 2, 28, NULL_1899), Some(60.0));
        assert_eq!(ymd_to_serial(1900, 3, 1, NULL_1899), Some(61.0));
        assert_eq!(ymd_to_serial(2024, 1, 1, NULL_1899), Some(45292.0));
        assert_eq!(serial_to_ymd(60.0, NULL_1899), Some((1900, 2, 28)));
        assert_eq!(serial_to_ymd(45292.0, NULL_1899), Some((2024, 1, 1)));
        assert_eq!(serial_to_ymd(0.0, NULL_1899), Some((1899, 12, 30)));
        assert_eq!(serial_to_ymd(-1.0, NULL_1899), Some((1899, 12, 29)));
    }

    #[test]
    fn known_serials_1904() {
        assert_eq!(ymd_to_serial(1904, 1, 1, NULL_1904), Some(0.0));
        assert_eq!(ymd_to_serial(2024, 1, 1, NULL_1904), Some(43830.0));
        assert_eq!(serial_to_ymd(43830.0, NULL_1904), Some((2024, 1, 1)));
    }

    #[test]
    fn the_three_null_dates_on_the_real_calendar() {
        let s1900 = DateSystem::Null1900;
        assert_eq!(ymd_to_serial(2024, 1, 1, s1900), Some(45290.0));
        assert_eq!(ymd_to_serial(1900, 1, 1, s1900), Some(0.0));
        assert_eq!(serial_to_ymd(0.0, s1900), Some((1900, 1, 1)));
        // Dates before the null date are negative serials, in every system.
        assert_eq!(ymd_to_serial(1900, 1, 1, NULL_1904), Some(-1460.0));
        assert_eq!(weekday_sun0(45290.0, s1900), 1);
    }

    #[test]
    fn overflow_and_weekday() {
        assert_eq!(
            ymd_to_serial(2024, 14, 1, NULL_1899),
            ymd_to_serial(2025, 2, 1, NULL_1899)
        );
        assert_eq!(
            ymd_to_serial(2024, 3, 0, NULL_1899),
            ymd_to_serial(2024, 2, 29, NULL_1899)
        );
        // 2024-01-01 was a Monday.
        assert_eq!(weekday_sun0(45292.0, NULL_1899), 1);
        assert_eq!(weekday_sun0(43830.0, NULL_1904), 1);
    }
}

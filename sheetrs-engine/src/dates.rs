//! Serial-number date arithmetic for both spreadsheet date systems.

use crate::options::DateSystem;

fn out_of_range<T>() -> Option<T> {
    crate::notes::note(crate::notes::DATE_RANGE);
    None
}

/// Days since 1970-01-01 of a proleptic Gregorian date (Hinnant's algorithm).
pub fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = y.div_euclid(400);
    let yoe = y.rem_euclid(400);
    let doy = (153 * (m + if m > 2 { -3 } else { 9 }) + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

pub fn civil_from_days(z: i64) -> (i64, i64, i64) {
    let z = z + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097);
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    (if m <= 2 { y + 1 } else { y }, m, d)
}

fn epoch(sys: DateSystem) -> i64 {
    sys.null_date
}

/// Serials between 0001-01-01 and 32767-12-31, the dates LibreOffice can write.
fn serial_range(sys: DateSystem) -> (i64, i64) {
    (
        days_from_civil(1, 1, 1) - epoch(sys),
        days_from_civil(32767, 12, 31) - epoch(sys) + 1,
    )
}

/// Serial number of a calendar date; months and days may overflow like in
/// `DATE()`. `None` when the result is outside the representable range.
pub fn ymd_to_serial(y: i64, m: i64, d: i64, sys: DateSystem) -> Option<f64> {
    ymd_to_serial_quiet(y, m, d, sys).or_else(out_of_range)
}

/// Like [`ymd_to_serial`], for text that merely may not be a date.
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

/// Like [`serial_to_ymd`], for values that are not known to be dates.
pub fn serial_to_ymd_quiet(serial: f64, sys: DateSystem) -> Option<(i64, i64, i64)> {
    let (min, max) = serial_range(sys);
    if !((min as f64)..(max as f64)).contains(&serial) {
        return None;
    }
    let s = serial.floor() as i64;
    Some(civil_from_days(epoch(sys) + s))
}

/// 0 = Sunday … 6 = Saturday.
pub fn weekday_sun0(serial: f64, sys: DateSystem) -> i64 {
    let s = serial.floor() as i64;
    // 1970-01-01 was a Thursday (4).
    (epoch(sys) + s + 4).rem_euclid(7)
}

pub fn unix_to_serial(secs: f64, sys: DateSystem) -> f64 {
    let days = secs / 86_400.0;
    days + (days_from_civil(1970, 1, 1) - epoch(sys)) as f64
}

pub fn is_leap(y: i64) -> bool {
    (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
}

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
    const NULL_1899: DateSystem = DateSystem::NULL_1899;
    const NULL_1904: DateSystem = DateSystem::NULL_1904;

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
    fn any_null_date_on_the_real_calendar() {
        let s1900 = DateSystem::NULL_1900;
        assert_eq!(ymd_to_serial(2024, 1, 1, s1900), Some(45290.0));
        assert_eq!(ymd_to_serial(1900, 1, 1, s1900), Some(0.0));
        let odd = DateSystem::from_ymd(1950, 6, 15).unwrap();
        assert_eq!(ymd_to_serial(2024, 1, 1, odd), Some(26863.0));
        assert_eq!(odd.ymd(), (1950, 6, 15));
        assert_eq!(DateSystem::from_ymd(2023, 2, 29), None);
        // Dates before the null date are negative serials, in every system.
        assert_eq!(ymd_to_serial(1900, 1, 1, NULL_1904), Some(-1460.0));
        assert_eq!(weekday_sun0(26863.0, odd), 1);
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

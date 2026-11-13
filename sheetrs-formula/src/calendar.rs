//! The calendar arithmetic of serial dates that the readers, the writers and the engine share.
//!
//! Days are counted from 1970-01-01 on the proleptic Gregorian calendar (Hinnant's algorithms); a
//! serial is a day count from the null date of a [`DateSystem`].

use crate::DateSystem;

/// Days since 1970-01-01 of a proleptic Gregorian date (Hinnant's algorithm).
#[must_use]
pub fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = y.div_euclid(400);
    let yoe = y.rem_euclid(400);
    let doy = (153 * (m + if m > 2 { -3 } else { 9 }) + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

/// The proleptic Gregorian date of a day number (days since 1970-01-01), as `(year, month, day)`.
#[must_use]
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

/// Serials between 0001-01-01 and 32767-12-31, the dates `LibreOffice` can write.
#[must_use]
pub fn serial_range(sys: DateSystem) -> (i64, i64) {
    let epoch = sys.days_from_unix_epoch();
    (
        days_from_civil(1, 1, 1) - epoch,
        days_from_civil(32767, 12, 31) - epoch + 1,
    )
}

/// The calendar date of a serial, for values that are not known to be dates: `None` outside the
/// range of [`serial_range`].
#[must_use]
pub fn serial_to_ymd_quiet(serial: f64, sys: DateSystem) -> Option<(i64, i64, i64)> {
    let (min, max) = serial_range(sys);
    if !((min as f64)..(max as f64)).contains(&serial) {
        return None;
    }
    let s = serial.floor() as i64;
    Some(civil_from_days(sys.days_from_unix_epoch() + s))
}

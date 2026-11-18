//! Serial numbers of calendar dates.

use crate::DateSystem;

/// Days from the serial zero date of the 1900 system (1899-12-30) back to 1970-01-01.
const EPOCH_TO_1900_ZERO: i64 = 25_569;
/// Days from the serial zero date of the 1904 system (1904-01-01) back to 1970-01-01.
const EPOCH_TO_1904_ZERO: i64 = 24_107;

impl DateSystem {
    /// The serial number of a calendar date of the proleptic Gregorian calendar,
    /// or `None` when the date does not exist.
    ///
    /// The 1900 system counts days from 1899-12-30, so 1900-03-01 is 61 and the
    /// non-existent 1900-02-29 is rejected. A year whose serial does not fit in 64 bits has none.
    #[must_use]
    pub fn serial_of_date(self, year: i64, month: u32, day: u32) -> Option<i64> {
        let epoch_days = days_from_civil(year, month, day)?;
        epoch_days.checked_add(match self {
            Self::Date1900 => EPOCH_TO_1900_ZERO,
            Self::Date1904 => EPOCH_TO_1904_ZERO,
        })
    }
}

fn is_leap_year(year: i64) -> bool {
    year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)
}

fn days_in_month(year: i64, month: u32) -> Option<u32> {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => Some(31),
        4 | 6 | 9 | 11 => Some(30),
        2 => Some(if is_leap_year(year) { 29 } else { 28 }),
        _ => None,
    }
}

/// Days from 1970-01-01 to the given date, counting backwards as negative; `None` when the
/// date does not exist or the count does not fit.
fn days_from_civil(year: i64, month: u32, day: u32) -> Option<i64> {
    if day == 0 || day > days_in_month(year, month)? {
        return None;
    }
    let shifted_year = if month <= 2 {
        year.checked_sub(1)?
    } else {
        year
    };
    let era = shifted_year.div_euclid(400);
    let year_of_era = shifted_year.rem_euclid(400);
    let shifted_month = i64::from(if month > 2 { month - 3 } else { month + 9 });
    let day_of_year = (153 * shifted_month + 2) / 5 + i64::from(day) - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    era.checked_mul(146_097)?
        .checked_add(day_of_era)?
        .checked_sub(719_468)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn serials_of_the_1900_system_count_from_1899_12_30() {
        let serial = |y, m, d| DateSystem::Date1900.serial_of_date(y, m, d);
        assert_eq!(serial(2020, 1, 2), Some(43_832));
        assert_eq!(serial(1899, 12, 31), Some(1));
        assert_eq!(serial(1900, 1, 1), Some(2));
        assert_eq!(serial(1900, 2, 28), Some(60));
        assert_eq!(serial(1900, 3, 1), Some(61));
        assert_eq!(serial(1, 1, 1), Some(-693_593));
        assert_eq!(serial(9999, 12, 31), Some(2_958_465));
    }

    #[test]
    fn serials_of_the_1904_system_count_from_1904_01_01() {
        let serial = |y, m, d| DateSystem::Date1904.serial_of_date(y, m, d);
        assert_eq!(serial(1904, 1, 1), Some(0));
        assert_eq!(serial(2020, 1, 2), Some(42_370));
        assert_eq!(serial(1900, 1, 1), Some(-1_460));
    }

    #[test]
    fn dates_that_do_not_exist_have_no_serial() {
        let serial = |y, m, d| DateSystem::Date1900.serial_of_date(y, m, d);
        assert_eq!(serial(2019, 2, 29), None);
        assert_eq!(serial(1900, 2, 29), None);
        assert_eq!(serial(2020, 2, 30), None);
        assert_eq!(serial(2020, 13, 1), None);
        assert_eq!(serial(2020, 0, 10), None);
        assert_eq!(serial(2020, 1, 0), None);
        assert_eq!(serial(i64::MAX, 1, 1), None);
        assert_eq!(serial(i64::MIN, 3, 1), None);
        assert!(serial(2020, 2, 29).is_some());
        assert!(serial(2000, 2, 29).is_some());
    }
}

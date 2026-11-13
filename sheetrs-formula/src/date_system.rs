//! The date system of a workbook: the date of serial 0.
//!
//! `LibreOffice` Calc offers three (in `Tools ▸ Options ▸ LibreOffice Calc ▸ Calculate`), and `SheetRS`
//! supports those three and no other:
//!
//! | Variant | Null date | In the dialog |
//! |---|---|---|
//! | [`DateSystem::Null1899`] | 1899-12-30 | `12/30/1899 (default)`. It is also what an XLSX in the 1900 system means |
//! | [`DateSystem::Null1900`] | 1900-01-01 | `01/01/1900 (StarCalc 1.0)` |
//! | [`DateSystem::Null1904`] | 1904-01-01 | `01/01/1904` (the 1904 system of Excel for Mac, `date1904` in an XLSX) |
//!
//! `LibreOffice` can store another null date, through its API or in a crafted file (`table:null-date`
//! of an ODS), and `SheetRS` gives [`UnsupportedDateSystem`] for it instead of numbering the dates
//! from a date it does not support: it is an approved difference (see
//! `docs/approved-differences.md`).

use std::fmt;

/// One of the three date systems of `LibreOffice` Calc.
///
/// The serials count days on the real calendar from the null date (no 1900-02-29, negative serials
/// allowed). The default is [`Null1899`](Self::Null1899).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub enum DateSystem {
    /// 1899-12-30, the default of `LibreOffice` and what an XLSX gets unless it says `date1904`.
    #[default]
    Null1899,
    /// 1900-01-01, the `StarCalc 1.0` option of `LibreOffice`.
    Null1900,
    /// 1904-01-01, the system of Excel for Mac (`date1904` of an XLSX).
    Null1904,
}

impl DateSystem {
    /// The three systems, the default first.
    pub const ALL: [Self; 3] = [Self::Null1899, Self::Null1900, Self::Null1904];

    /// The null date, the date of serial 0, as `(year, month, day)`.
    #[must_use]
    pub const fn null_date(self) -> (i64, u32, u32) {
        match self {
            Self::Null1899 => (1899, 12, 30),
            Self::Null1900 => (1900, 1, 1),
            Self::Null1904 => (1904, 1, 1),
        }
    }

    /// Days from 1970-01-01 to the null date (negative: the null dates are before it).
    #[must_use]
    pub const fn days_from_unix_epoch(self) -> i64 {
        match self {
            Self::Null1899 => -25_569,
            Self::Null1900 => -25_567,
            Self::Null1904 => -24_107,
        }
    }

    /// Days from the default null date, 1899-12-30, to this one: 0, 2 or 1462. A date numbered from
    /// the default system has this many more days than from this one.
    #[must_use]
    pub const fn days_from_default(self) -> i64 {
        self.days_from_unix_epoch() - Self::Null1899.days_from_unix_epoch()
    }

    /// The system whose null date is `year-month-day`.
    ///
    /// # Errors
    ///
    /// [`UnsupportedDateSystem`] when it is not one of the three, whatever the date is.
    pub fn from_null_date(year: i64, month: u32, day: u32) -> Result<Self, UnsupportedDateSystem> {
        Self::ALL
            .into_iter()
            .find(|system| system.null_date() == (year, month, day))
            .ok_or(UnsupportedDateSystem { year, month, day })
    }

    /// The system named by an ISO 8601 date, `YYYY-MM-DD` (what `table:null-date` holds, and what
    /// [`Display`](fmt::Display) writes).
    ///
    /// # Errors
    ///
    /// [`UnsupportedDateSystem`] when it is not the null date of one of the three. `None` when the
    /// text is not a `YYYY-MM-DD` date at all.
    #[must_use]
    pub fn from_iso(text: &str) -> Option<Result<Self, UnsupportedDateSystem>> {
        let mut parts = text.trim().splitn(3, '-');
        let year = parts.next()?.parse::<i64>().ok()?;
        let month = parts.next()?.parse::<u32>().ok()?;
        let day = parts.next()?.parse::<u32>().ok()?;
        Some(Self::from_null_date(year, month, day))
    }
}

impl fmt::Display for DateSystem {
    /// The null date, `YYYY-MM-DD`.
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let (year, month, day) = self.null_date();
        write!(f, "{year:04}-{month:02}-{day:02}")
    }
}

/// The text is none of the names of the three date systems (`1899`, `1900`, `1904`, or the null date
/// as `YYYY-MM-DD`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParseDateSystemError(String);

impl fmt::Display for ParseDateSystemError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "must be 1899 (1899-12-30, the default), 1900 (1900-01-01) or 1904 (1904-01-01), the \
             three date systems of LibreOffice, got `{}`",
            self.0
        )
    }
}

impl std::error::Error for ParseDateSystemError {}

impl std::str::FromStr for DateSystem {
    type Err = ParseDateSystemError;

    /// `1899` or `1899-12-30`, `1900` or `1900-01-01`, `1904` or `1904-01-01`.
    fn from_str(text: &str) -> Result<Self, Self::Err> {
        match text {
            "1904" | "1904-01-01" => Ok(Self::Null1904),
            "1900" | "1900-01-01" => Ok(Self::Null1900),
            "1899" | "1899-12-30" => Ok(Self::Null1899),
            _ => Err(ParseDateSystemError(text.to_string())),
        }
    }
}

/// A date system other than the three `SheetRS` supports: the workbook's date of serial 0 is not
/// 1899-12-30, 1900-01-01 or 1904-01-01.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct UnsupportedDateSystem {
    /// The null date the workbook declares.
    pub year: i64,
    /// Its month, 1 to 12.
    pub month: u32,
    /// Its day of the month.
    pub day: u32,
}

impl fmt::Display for UnsupportedDateSystem {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "unsupported date system: the workbook numbers its dates from {:04}-{:02}-{:02}; the \
             supported ones are 1899-12-30 (the default), 1900-01-01 and 1904-01-01",
            self.year, self.month, self.day
        )
    }
}

impl std::error::Error for UnsupportedDateSystem {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_three_systems_and_no_other() {
        for system in DateSystem::ALL {
            let (y, m, d) = system.null_date();
            assert_eq!(DateSystem::from_null_date(y, m, d), Ok(system));
            assert_eq!(DateSystem::from_iso(&system.to_string()), Some(Ok(system)));
        }
        let other = DateSystem::from_null_date(1950, 6, 15).unwrap_err();
        assert_eq!((other.year, other.month, other.day), (1950, 6, 15));
        let text = other.to_string();
        assert!(text.contains("1950-06-15"), "{text}");
        for supported in ["1899-12-30", "1900-01-01", "1904-01-01"] {
            assert!(text.contains(supported), "{text}");
        }
        // Not even a calendar date is accepted.
        assert!(DateSystem::from_null_date(1904, 2, 30).is_err());
        assert_eq!(DateSystem::from_iso("soon"), None);
    }

    #[test]
    fn the_offsets_are_the_ones_of_the_calendar() {
        assert_eq!(DateSystem::Null1899.days_from_default(), 0);
        assert_eq!(DateSystem::Null1900.days_from_default(), 2);
        assert_eq!(DateSystem::Null1904.days_from_default(), 1462);
        assert_eq!(DateSystem::default(), DateSystem::Null1899);
    }
}

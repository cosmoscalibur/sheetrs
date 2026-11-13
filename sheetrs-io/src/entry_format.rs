//! The number format a typed text gets, shared by the text context of a query and the writers.

/// The number format `LibreOffice` gives to a cell in `General` that is typed a text it reads as
/// something other than a plain number (`50%` a percentage, `03/15/2024` a date, `TRUE` a boolean).
///
/// The formats are those of the `en-US` locale, the only one the text context reads.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum EntryFormat {
    /// `0.00%`.
    Percent,
    /// The standard date of the locale, `MM/DD/YY`.
    Date,
    /// `YYYY-MM-DD`, for a date typed in that shape.
    DateIso,
    /// `MM/DD/YY HH:MM AM/PM`: a date and a time of day.
    DateTime,
    /// `YYYY-MM-DD HH:MM:SS`.
    DateTimeIso,
    /// `YYYY-MM-DD"T"HH:MM:SS`.
    DateTimeIsoT,
    /// `HH:MM:SS AM/PM`.
    Time,
    /// `[HH]:MM:SS`: a time of 24 hours or more.
    TimeTotal,
    /// `[HH]:MM:SS.00`: a time with a fraction of a second.
    TimeTotalFraction,
    /// `0.00E+00`.
    Scientific,
    /// `[$$-409]#,##0.00;[RED]-[$$-409]#,##0.00`.
    Currency,
    /// `# ?/?`.
    Fraction,
    /// `BOOLEAN`.
    Boolean,
}

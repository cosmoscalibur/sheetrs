### Fixed

- The text a cell shows for a date, a time or an elapsed time is what LibreOffice writes where it differed: dates before 1582-10-15 in the Julian calendar (and the years before 1 as -1, -2…), `#FMT` for a date format outside the range of its number formatter (a clock format writes the time of day of any number), `[hh]`/`[mm]`/`[ss]` with their digits, an elapsed time past 2^32 seconds as `#FMT`, the seconds cut and the fractions rounded without reaching the next second (`23:59:59.99`), and a year under 10 without a zero in two digits. `TEXT` takes the calendar, the digits and the rounding too.
- The reader of an ODS reads a `office:date-value` before the year 1 (`-0001-12-31`, 1 BC) as the date it is, where it was text.

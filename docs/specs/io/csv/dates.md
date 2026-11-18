# Dates in delimited text

## Observations

| Behaviour | Test |
| --- | --- |
| A field `YYYY-MM-DD`, optionally followed by `T` or `t` and `hh:mm:ss` with an optional fraction after `.` or `,`, between ASCII spaces, is a serial number even when special numbers are not detected. The year has four or five digits, an optional minus sign and a value from -32767 to 32767 other than 0. There are no offsets, and the hour has two digits. | `iso_dates_over_csv` in `sheetrs-csv` |
| Dates up to 1582-10-04 are dates of the Julian calendar and the days run on to 1582-10-15, so 5 to 14 October 1582 do not exist; a year before 1 counts as in the calendar without a year 0. Whether February has 29 days follows the Gregorian rule in both calendars. | `iso_dates_over_csv` in `sheetrs-csv` |
| Hour 24 stands for midnight at the end of the day and takes no minutes, seconds or fraction. | `iso_dates_over_csv` in `sheetrs-csv` |

## Open defects

None known.

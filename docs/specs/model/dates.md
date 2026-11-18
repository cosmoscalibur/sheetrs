# Dates in the model

Serial day numbers shared by every format.

## Observations

| Behaviour | Test |
| --- | --- |
| A serial number counts days from 1899-12-30, or from 1904-01-01 in the 1904 system; the zero can be any date in a format that lets the file choose it. | `date_cells_follow_the_date_system_flag_over_xlsx` in `sheetrs-xlsx` and `dates_follow_the_null_date_over_ods` in `sheetrs-ods` |
| The days between two dates are the days of the proleptic Gregorian calendar, with no 29 February 1900 to skip. | `dates_follow_the_null_date_over_ods` in `sheetrs-ods` |

## Open defects

- Before year 1 the formats disagree: the dates of an XLSX cell of type `d` count as astronomical
  years, an ODS `office:date-value` as years without a year 0, and delimited text as Julian dates
  without a year 0. The reader of XLSX differs from LibreOffice by one day for some years
  (`-4712-01-01`, `-32768-01-01`, `-2020-01-02`, `-0004-12-31`, `-0005-01-01`, `-0400-12-31`).

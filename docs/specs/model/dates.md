# Dates in the model

Serial day numbers shared by every format.

## Observations

| Behaviour | Test |
| --- | --- |
| A serial number counts days from 1899-12-30, or from 1904-01-01 in the 1904 system. | `date_cells_follow_the_date_system_flag_over_xlsx` in `sheetrs-xlsx` |

## Open defects

- The day count of dates before year 1 differs from LibreOffice's by one day for some
  years (`-4712-01-01`, `-32768-01-01`, `-2020-01-02`); the model counts the proleptic
  Gregorian calendar.

### Added

- `dates` option of a query (`sheetrs query --dates`, `"dates": true` in the JSON query and in `query_workbook`, `dates=True` in Python, `dates: bool` in `RangeOptions`, `FindOptions` and `Query`): by default every number is the serial the file stores; with the flag a number whose format is a date, a time or both is given as `{"type": "date" | "time" | "datetime", "value": "<ISO 8601>"}` (`sheetrs_protocol::DateValue`), worked out with the date system of the file (the usual one, 1904, or the `table:null-date` of an ODS). The conversion follows LibreOffice (checked against its own display in the three date systems: seconds cut, fractions rounded, Julian calendar before 1582-10-15, no 1900 leap-year day); a serial whose year is outside 1 to 9999, an elapsed time and any other number stay numbers.
- `Source::date_system` (`sheetrs_io`): the date system of a workbook, a `DateSystem` (`sheetrs_formula`) of the three of LibreOffice Calc: `Null1899` (1899-12-30, the default and what an XLSX gets unless it says `date1904`), `Null1900` (1900-01-01, StarCalc 1.0) and `Null1904` (1904-01-01).

### Fixed

- `display` and a search by formatted text numbered the dates of an ODS whose null date is not the default from 1899-12-30; they use the date system of the file now.

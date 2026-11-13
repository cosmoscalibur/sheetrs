### Changed

- The date system is one of the three of the LibreOffice Calc interface, and no other: `DateSystem` (`sheetrs_formula`) is `Null1899` (1899-12-30, the default, and what an XLSX gets unless it says `date1904`), `Null1900` (1900-01-01, StarCalc 1.0) and `Null1904` (1904-01-01). It replaces the number of days of the null date in the options of the engine, the reader (`Workbook::date_system`, `DocumentSettings::date_system`, `Source::date_system`), the text context of the query (`TextContext::new(locale, date_system)`), the editor and the compiler. `--date-system` accepts `1899`, `1900` or `1904` (or their null dates).
- An ODS whose `table:null-date` is none of the three (LibreOffice can store one through its API) is an error that names the date and the three supported ones (`sheetrs_io::Error::UnsupportedDateSystem`, `sheetrs.UnsupportedDateSystemError` in Python), raised when the workbook is read, by every tool, instead of numbering the dates from it. It is an approved difference (`docs/approved-differences.md`).
- `sheetrs edit --replace` types a date in the date system of the workbook, and the text an edited ODS cell shows is written in it (it was written in the default one in a workbook of 1904).

### Removed

- The support of an arbitrary null date, and its fixtures (`null_date_1950.ods`, `dates_1950.ods`): `sheetrs-io/tests/assets/other_null_date.ods` is the file the readers refuse, and `dates_1900.ods` (judged by LibreOffice) covers the 1900-01-01 system.

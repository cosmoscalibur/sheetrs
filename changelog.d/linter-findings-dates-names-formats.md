### Fixed

- ODS: a defined name that `LibreOffice` writes after the last table (for an ODS that links an external workbook, the cache of that workbook) no longer takes the table as its scope; REF311 reports the workbook and the sheet.
- The last-modified date of an ODS saved by `LibreOffice` (`dc:date` without a time zone and with a fraction of a second) is read, taken as UTC, so FILE1002 reports old ODS files; a date without a zone is accepted in an XLSX too.
- DATA703 compares a date format without its locale mark (`[$-409]mm/dd/yyyy` equals `mm/dd/yyyy`), as the rule always did for backslashes.

### Changed

- The 1904 date system (FILE1003) is covered by a unit test of the rule on a read-result model and by `sheetrs-io/tests/date_system_tests.rs`, which makes small workbooks in memory; the extra books `lint_book_1904.*` are gone.

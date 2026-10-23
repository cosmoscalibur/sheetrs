### Fixed

- ODS dates (`office:date-value`) are now numbered from the document's `table:null-date` (1899-12-30 by default, as LibreOffice does) instead of from 1899-12-31. The value is the number the spreadsheet itself uses, so it equals the XLSX serial for dates from 1900-03-01 on; before, every ODS date was one lower.
- ODS time cells (`office:time-value`, e.g. `PT13H30M00S`) are read as a fraction of a day instead of their display text.

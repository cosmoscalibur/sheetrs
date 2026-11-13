### Added

- `Find.filtered`, `sheetrs query --filtered` and `sheetrs edit --filtered` (`SearchFiltered` of LibreOffice): a search or replacement now leaves out the rows an autofilter hides unless it asks for them; rows hidden by hand are always reached. The selective reader reports the filtered rows (`CellSink::filtered_rows`, `Selection::with_filtered_rows`): an ODS marks them (`table:visibility="filter"`), an XLSX counts the rows hidden inside the `<autoFilter ref>` of the sheet, as LibreOffice does when it opens the file. Tested in both formats against workbooks and `findAll`/`replaceAll` results made by LibreOffice.
- `Source::defined_names`: the defined names of a workbook (and those defined in more than one scope) without reading the cells of an XLSX.
- The format a formula cell takes from its result (`sheetrs query`, `display` and `number_format`) is now worked out for `IF`, `CHOOSE` and `INDIRECT` whose condition, index or text is a function of constants, cells and defined names (`ISNUMBER`, `LEN`, `ROW`, `CONCATENATE`, `SUM`, `COUNT`, …), and through defined names that stand for a cell, a range or a constant; fixture `inherit_more.*` and its oracle.

### Changed

- A search in XLSX threaded comments is documented and tested as LibreOffice does it: it reads the legacy note Excel writes for the thread and ignores the threaded part.

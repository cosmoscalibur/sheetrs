### Fixed

- The shared formulas of an XLSX shift whole columns and rows (`SUM(B:B)` in the next column is `SUM(C:C)`), and leave alone what is not a cell of the sheet: names that look like a cell (`TOTALSALES2024`, `ABCD1`, `A0`), calls (`LOG10(`, `DEC2BIN(`), text between quotes and the names inside structured references. They overflowed a column index for a token of seven or more capital letters and a digit.
- A shared formula whose cell would read outside the sheet is `Sheet!#REF!` there, as LibreOffice computes it: the workbook builds, the cell is `#REF!` when it runs, and the linter reports it as an error cell. `Sheet!#REF!` is `#REF!` in the compiler when the sheet exists.
- CALC202 and REF310 read a token that is not a cell of the sheet (`A0`, `XFE1`, `TOTALSALES2024`) as a name, and CALC205 no longer reads a nested call such as `ABS(A1)` as a range (#69).

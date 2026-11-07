### Fixed

- `ods`: a reference to another sheet written without `$` (`[T.A1]`, `SUM([T.A1:.A3])`) and a
  quoted sheet name with spaces, dots or apostrophes (`['Q1.2024'.B2]`, `['O''Brien'.A2]`) are now
  normalized to the formula an XLSX holds for the same cell (`T!A1`, `SUM(T!A1:A3)`,
  `'Q1.2024'!B2`, `'O''Brien'!A2`); before, 20 of 41 formulas of a document made by
  `LibreOffice` differed from the XLSX it exports. Text inside string literals is left as it is.
- `ods`: the number format of a cell whose style holds only a data style (a self-closing
  `style:style`, which `LibreOffice` writes for dates) is read; those cells had no format and a
  date read as a number.

### Added

- `ods`: the number formats of a cell are read for every family of data style (number, grouping,
  percentage, scientific, fraction, time, duration, currency, boolean, text, and the sections with
  colours and conditions of `style:map`), written as the format code the XLSX of the same cell
  holds (`0.00%`, `[hh]:mm:ss`, `#,##0.00 [$EUR]`, `0.00;[RED]-0.00`), and a cell style without a data
  style takes the one of its parent. 38 formats of `LibreOffice` read the same in both formats; the
  one that differs is the built-in `0`, which `LibreOffice` writes to an ODS as no style (General).
  The reading of the styles lives in one module instead of three copies.

### Changed

- `ods`: an error that only `LibreOffice` has (`Err:502` invalid argument, `Err:522` circular
  reference, `Err:523` no convergence, …) keeps its code in the cell (`ExcelError::Libre(code)`,
  written `Err:522`) instead of reading as `#VALUE!`, which is what `LibreOffice` shows in an ODS
  (it writes `#VALUE!` for them to an XLSX, so there the cell is `#VALUE!`). The codes that have an
  Excel name (`Err:524` is `#REF!`, `Err:525` `#NAME?`, `Err:532` `#DIV/0!`, `Err:503` `#NUM!`,
  `Err:519` `#VALUE!`) read as that. `sheetc` builds a circular reference of an ODS as `#CIRC!`
  again, as its loader intended.

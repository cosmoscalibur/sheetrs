### Fixed

- `ods`: a reference to another sheet written without `$` (`[T.A1]`, `SUM([T.A1:.A3])`) and a
  quoted sheet name with spaces, dots or apostrophes (`['Q1.2024'.B2]`, `['O''Brien'.A2]`) are now
  normalized to the formula an XLSX holds for the same cell (`T!A1`, `SUM(T!A1:A3)`,
  `'Q1.2024'!B2`, `'O''Brien'!A2`); before, 20 of 41 formulas of a document made by
  `LibreOffice` differed from the XLSX it exports. Text inside string literals is left as it is.
- `ods`: the number format of a cell whose style holds only a data style (a self-closing
  `style:style`, which `LibreOffice` writes for dates) is read; those cells had no format and a
  date read as a number.

### Known differences

- `ods`: time, duration, percentage and currency formats are not read yet (the XLSX gives them).
- `ods`: a `LibreOffice` error that Excel does not have (`Err:502`, `Err:522`) reads as `#VALUE!`,
  as `LibreOffice` itself writes it to an XLSX.

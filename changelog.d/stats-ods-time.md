### Changed

- `sheetrs stats` on an ODS no longer converts every formula to Excel's notation, which the counts do not use:
  2.4 M formulas take 2.9 s instead of 5.4 s, with the same output. The XLSX read is unchanged.

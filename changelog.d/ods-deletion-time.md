### Fixed

- Fix `sheetrs edit --delete-columns` on an ODS taking five times as long as on the XLSX of the same workbook (136-160 s against 22 s, 2.4 M formulas): the empty run that closes each row up to the last column of the sheet was written shorter by the columns deleted, so the next read of the file (the recalculation of the plan) took it for formatted cells and built about 16 000 of them per row. The run keeps reaching the last column.

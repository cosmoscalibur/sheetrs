### Changed

- `sheetrs-compiler`, `sheetrs-formula`: a reference over several sheets handed through `IF` or `CHOOSE` is read as the cells of every sheet (`SUM(IF(TRUE,Data:Wide!A1:A3,0))` is 39 as in LibreOffice); the first sheet quoted alone (`'Data':'Wide'!A1`, `'Data':Wide!A1`), which stopped the build, and the last (`Data:'Wide'!A1`, `#REF!` before) are read as `Data:Wide!A1`; a column with no row on the last sheet is on the row of the first corner (`Data!A2:Wide!A`). Rows of the oracle for them.

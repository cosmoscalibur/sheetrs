### Changed

- `sheetc`: without declared inputs, only the constants that start a path of the dependency graph (a formula reads them, or a defined name or table covers them) are inputs; an isolated constant is not (it can still be an output). A workbook that resolves references at run time (`INDIRECT`, `OFFSET`) keeps every constant as an input.
- The engine carries the identity of LibreOffice's own errors (`Err:502`, `Err:504`, `Err:523`…) through the results: JSON gives `{"error": "#VALUE!", "code": "Err:502"}`, a recalculated ODS writes the code and an XLSX the closest error it can hold. Exact where LibreOffice answers one code per function; the others answer the most frequent code.

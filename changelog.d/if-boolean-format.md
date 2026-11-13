### Fixed

- A formula whose result is what `IF` or `CHOOSE` chose element by element is a boolean when the last branch that ran set a boolean format, as in LibreOffice: `IF({1;0},TRUE())` is TRUE (it was the number 1), `IF({1;0},0,TRUE())` FALSE, `IF({0;1},TRUE(),5)` TRUE; a branch that gives a number through a function or an operator (`SUM(1)`, `DATE(…)`, `5+1`) leaves a number, and what a function takes of the array (`INDEX`, `ISLOGICAL`, `IFERROR`) is still the numbers. The cases and the ones left are in `sheetrs-compiler/docs/FIDELITY.md`.
- The reader of an ODS takes an `office:boolean-value` that is a number (LibreOffice writes `5` for a cell that shows TRUE and holds 5) as TRUE unless it is zero, as LibreOffice reads it back; it was FALSE.

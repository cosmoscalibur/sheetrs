### Added

- `sheetc run FILE --outputs any ... --write-workbook OUT`: the answer is written into a copy of the
  workbook instead of printed. Formulas, styles and number formats stay, every formula holds the
  result of the run and the inputs given are written in their cells; the output is the whole
  answer, in the format of the template, in the formats of its cells (so a date comes out as a
  date, a percentage as a percentage). `--template PATH` uses another workbook of the same layout.
  Errors are written as the format has them (`#CIRC!` is `Err:522` in an ODS). The program must be
  built with `outputs = any`; the cells of an array formula other than its first keep the results
  the template cached.
- `sheetc build FILE --embed-template`: the workbook travels inside the executable (made from `sheetc`
  itself) as the template; `./model --write-workbook OUT` writes the answer and `--template PATH`
  replaces the embedded one. The payload has a second trailer version (`SHEETRS2`) that also holds
  the template; executables without one are unchanged.
- `sheetrs-edit`: `Editor::write_cells(CellWrites)` writes values decided by the caller, keeping the
  formula, array and style of a cell for `CellWrites::result` (the result a program computed) and
  replacing the constant for `CellWrites::constant`; `fill_template` runs a program and writes its
  answer into a template. `Report::written` counts the cells. A write is a pass of its own and does
  not mark the workbook to be recalculated.

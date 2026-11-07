### Added

- `sheetc build FILE --embed-template` and `sheetc run FILE --embed-template`: the workbook is its own
  template and the answer is the workbook. The executable (made from `sheetc` itself) carries the
  workbook and gives no JSON: `--output OUT` writes a copy with the result of every formula and the
  inputs given in their cells, in the formats of the cells; `--meta` and `--info` are answered as
  before. The outputs the workbook declares are discarded, its inputs are what is accepted.
  Recalculating a workbook is the case with no inputs and needs `--recalculate`. There is no
  external template, and an executable built without `--embed-template` has no capability to write
  workbooks. The payload has a second trailer version (`SHEETRS2`) that also holds the workbook;
  executables without one are unchanged.
- `sheetrs-edit`: `Editor::write_cells(CellWrites)` writes values decided by the caller, keeping the
  formula, array and style of a cell for `CellWrites::result` (the result a program computed) and
  replacing the constant for `CellWrites::constant`; `fill_template` runs a program and writes its
  answer into a template. `Report::written` counts the cells. A write is a pass of its own and does
  not mark the workbook to be recalculated.

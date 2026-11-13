### Added

- `sheetc build FILE --embed-template` and `sheetc calculate FILE --embed-template`: the workbook is its own
  template and the answer is the workbook. The executable (made from `sheetc` itself) carries the
  workbook: `--output OUT` writes a copy with the results of the outputs that are formulas and the inputs given in
  their cells, in the formats of the cells; `--meta` and `--info` are answered as before. With
  nothing declared the formula cells are the outputs and the constants a path of the dependencies starts
  from the inputs (a range, or a defined name inside them), and what the settings declare narrows them (see the entry on JSON in
  the embedded executable). There is no external template, and an executable built without
  `--embed-template` has no capability to write workbooks (its default stub is still `sheetrs-rt`).
  The payload has a second trailer version (`SHEETRS2`) that also holds the workbook; executables
  without one are unchanged.
- `sheetc build|calculate FILE --recalculate --output OUT`: a mode of its own, the workbook recalculated as
  it is. It takes no input (a request with inputs is an error), gives no JSON, ignores what the
  settings declare and cannot be combined with `--embed-template`.
- `sheetrs-engine`: `ir::Mode` (`Values`, `Embedded`, `Recalculate`) in the program; `--meta` of an
  executable says it.
- `sheetrs-edit`: `Editor::write_cells(CellWrites)` writes values decided by the caller, keeping the
  formula, array and style of a cell for `CellWrites::result` (the result a program computed) and
  replacing the constant for `CellWrites::constant`; `sheetrs_answer::fill_template` runs a program and writes its
  answer into a template. `Report::written` counts the cells. A write is a pass of its own and does
  not mark the workbook to be recalculated.

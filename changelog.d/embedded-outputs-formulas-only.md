### Changed

- `sheetc --embed-template`: with nothing declared, the outputs of an embedded workbook are its formula
  cells and no constant, as the rectangles that hold exactly them (`--meta` lists them; with more than
  4 096 separate rectangles, the used area of the formulas of each sheet, with a note). The constants
  are already part of the template, so reading and carrying them as outputs was cost with no use.
  `--recalculate` and the programs that give values are unchanged. `--info` says `formulas` for the
  outputs option of an embedded program.
- `sheetrs_answer::fill_template` asks the program for its output rectangles and writes the inputs from
  what the request gave, not from the answer.

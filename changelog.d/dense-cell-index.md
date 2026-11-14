### Changed

- The engine finds the slot of a cell in a table per sheet and row instead of a hash map: a run of a
  compiled workbook of 800 000 formulas takes 0.56 s instead of 0.74 s, with the same results.

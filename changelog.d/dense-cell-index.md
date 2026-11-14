### Changed

- The engine finds the slot of a cell in a table per sheet and row instead of a hash map, finds a
  constant by its row inside a decoded block, and keeps the index of a program for the engines that run
  it after the first. On a workbook of 2.4 M formulas a compiled executable answers a request in 1.69 s
  instead of 2.50 s, and `--serve` takes 0.65 s per request instead of 1.5 s, with the same results.
- Writing the workbook of a calculation no longer copies the map of cells to write: a recalculation of
  2.4 M formulas takes 19 % less time and 12 % less peak memory with the same file.

### Fixed

- Fix `sheetrs edit --delete-columns` on an ODS taking five times as long as on the XLSX of the same workbook (136-139 s against 22 s, 2.4 M formulas): the empty run that closes each row up to the last column of the sheet was written shorter by the columns deleted, so the next read of the file (the recalculation of the plan) took it for formatted cells and built about 16 000 of them per row. The run keeps reaching the last column. The same file as `sheetrs edit` writes it now takes 55 s, and the output is byte-identical to the one of the previous binary, apart from those runs.

### Changed

- The analysis of the stale formulas of a deletion takes the cached values it needs from the model the compiler already loads, instead of reading the ODS a second time, and the check of an ODS for formulas with no cached result looks at the bytes of the cell tags before it builds XML events (the reader still decides when a cell may hold its result in a paragraph).
- `sheetrs_compiler::memory::compile_keeping`: `compile` that also hands a closure the model before it is compiled.

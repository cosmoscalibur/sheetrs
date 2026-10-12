### Added

- Add per-sheet used cells, errors and empty cells to `sheetstats`, with workbook totals for every per-sheet metric.
- Add `Sheet::xml_size_bytes`, the uncompressed size of the XML describing a sheet, for XLSX and ODS.

### Changed

- `sheetstats` groups its output by sheet. JSON output moves per-sheet metrics into a `sheets` array (replacing `sheet_sizes`, `formula_stats` and `cell_stats`).
- `sheetstats` reports the uncompressed XML size of each sheet for both XLSX and ODS, instead of the compressed size for XLSX only. Sheet sizes now follow the workbook's sheet relationships rather than the `sheetN.xml` file number.
- `sheetstats` counts as used only cells with a value or a formula; cells kept only for their style are no longer counted.

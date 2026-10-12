### Added

- Add `sheetcli --convert-to-values` to replace formulas by their cached values in whole sheets (`Sheet`), columns (`Sheet!B:D`) or rows (`Sheet!3:10`) of XLSX workbooks (#11). Shared formula cells left outside the scope get their own formula, and the calculation chain is removed.
- Add `writer::ValueScope` and `WorkbookModifications::convert_to_values` to the library.

### Changed

- `WorkbookModifications` gains the `convert_to_values` field; struct literals must add it or use `..Default::default()`.

### Fixed

- Fix XLSX error constants (`t="e"` cells without a formula) being read as empty cells. ERR102 now reports them.
- Fix XLSX sheets whose relationship target is an absolute part name (`/xl/worksheets/sheet1.xml`) not being found by name.

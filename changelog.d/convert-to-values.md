### Added

- Add `sheetcli --convert-to-values` to replace formulas by their cached values in whole sheets (`Sheet`), columns (`Sheet!B:D`) or rows (`Sheet!3:10`) of XLSX and ODS workbooks (#11). In XLSX, shared formula cells left outside the scope get their own formula, and the calculation chain is removed.
- Add ODS support to `sheetcli` modifications (`--remove-sheets`, `--remove-ranges`, `--convert-to-values`) and `writer::modify_workbook_ods`.
- Add `writer::ValueScope` and `WorkbookModifications::convert_to_values` to the library.

### Changed

- `WorkbookModifications` gains the `convert_to_values` field; struct literals must add it or use `..Default::default()`.
- Removing every sheet of a workbook is an error.
- XLSX modifications stream the converted worksheets and copy unchanged parts as they are, and ODS modifications stream `content.xml`, keeping at most one row of a sheet in memory: peak memory on the production files drops from up to 366 MB (XLSX) and 1.2 GB (ODS) to under 10 MB.

### Fixed

- Fix modifying a workbook with its input file as output truncating the input before reading it; it is now an error.
- Fix XLSX error constants (`t="e"` cells without a formula) being read as empty cells. ERR102 now reports them.
- Fix XLSX sheets whose relationship target is an absolute part name (`/xl/worksheets/sheet1.xml`) not being found by name.
- Fix `--remove-sheets` on XLSX deleting the worksheet part and relationship named after the `sheetId` instead of those of the sheet, and leaving sheet-local names, `localSheetId` and workbook view indices pointing at the wrong sheets. The calculation chain is now removed along with sheets.

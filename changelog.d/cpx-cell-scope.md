### Changed

- Report CPX504–CPX509 with cell scope instead of sheet scope, so text output
  lists them under their formula cell and JSON exposes the cell in `scope`.
  Their messages no longer end with ` at <cell>`, and their violation data
  structs (`DeepIfNestingData`, `DeepFormulaNestingData`, `ManyOperationsData`,
  `ManyReferencesData`, `LongFormulaData`, `MultipleSheetRefData`) drop the
  `cell` field.

### Fixed

- Fix the text output of `sheetlint` and the WASM build listing the cells of a sheet in string order (`A1`, `A10`, `A2`) instead of row and column order.

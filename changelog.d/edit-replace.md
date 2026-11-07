### Added

- `sheetrs-edit`: `Editor::replace(scopes, Replace)` and `sheetedit --replace TEXT --with TEXT
  [--scope SCOPE] [--regex | --wildcards] [--match-case] [--entire-cells] [--in value|formula]
  [--include-formulas]`: Find & Replace over the workbook, a sheet or a range, as `LibreOffice` does
  it (observed in 24.2.7; the reference version 26.8 could not be installed), in XLSX and ODS. The
  text a search sees is the one of the input line of the cell, every hit of a cell is replaced, the
  replacement is entered as if it were typed (`12` a number, `=1+1` a formula, nothing an empty
  cell), a formula is searched and replaced in the notation of `LibreOffice` and written back in the
  notation of the file (Excel's, `OpenFormula`'s), and a shared formula group whose first cell is
  replaced gets a formula of its own per cell. `save` and `overwrite` return a `Report` with the
  cells replaced and the ones found and left as they were (a replacement that leaves something that
  is not a formula, or names a missing sheet, is not applied; `LibreOffice` keeps an error cell).
  Searching comments is not supported. `tests/replace.rs` compares the result with the workbook
  `LibreOffice` leaves after the same `replaceAll`, for 26 cases in both formats.
- `sheetrs-formula`: `calc_read::from_calc` reads a formula typed in the notation of `LibreOffice`
  (`$'My Sheet'.A1:B2`, `;` or `,` between arguments) and writes it in Excel's or `OpenFormula`'s
  notation; it is the inverse of `calc_print::to_calc`.
- `sheetrs-query`: `text::TextContext::typed` (what a text becomes when typed into a cell).

### Changed

- `sheetrs-query`: with `entire_cells` the **first** match must cover the whole text (`a|ab` does
  not match `ab`), and the wildcard `*` takes as little as it can, as `LibreOffice` does.
- `sheetrs-edit`: `Editor::save` and `Editor::overwrite` return a `Report` (empty unless the plan
  replaces text).

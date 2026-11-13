### Added

- `Find::formatted` (`SearchFormatted` of LibreOffice): search the text each cell shows (`1,234.50`, `15 Mar 2023`) instead of its input line. In `sheetrs-query` (`--formatted`, the searches are compared with `findAll` of LibreOffice in `tests/libreoffice_find.rs`) and in `sheetrs-edit` (`sheetrs edit --replace … --formatted`, with fixtures made by LibreOffice's own `replaceAll`).
- Typing into a cell takes the currency symbol of the locale (`$20.00`, `-$5`, `($5)`, `5$`) and the literal text of the cell's format (`x9` in a cell of format `"x"0` is `9`) as LibreOffice does.

### Fixed

- A search by value over the results of formulas (`include_formulas`) looks at the text the result has with the format the cell takes from it (a date for `=A1+1` over a date), as LibreOffice does.

### Known gaps

- `SearchFiltered`: LibreOffice skips rows hidden by an autofilter unless asked; SheetRS searches them (pending).

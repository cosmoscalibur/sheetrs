### Added

- `Book::recalculation_marked` reads the mark to calculate every formula when a workbook is opened (`fullCalcOnLoad` of an XLSX; an ODS has none) in one place: `sheetc build` and `calculate`, `BookCompile` and `sheetrs edit` take a marked workbook as a run with `calculate_all`, and the workbook a calculation writes leaves without the mark.
- `Options::request_outputs` and `Spec::whole_workbook`: a calculation declares what it is asked and nothing more. With the workbook as the answer (the embedded mode) it declares only its inputs and every formula comes out calculated, the ones the workbook declares no output for included; with no workbook, the `ranges` of the request are the declared outputs of the build when the workbook declares none, and the program keeps only their path.
- `sheetrs-edit/tests/shared_beyond.rs` and its oracle (`make_shared_oracle.py`): the copies of a shared formula that would read outside the sheet, against what LibreOffice shows of the file it saved and opened again.

### Changed

- A range that a deletion leaves with one cell is that cell in the text of an XLSX (`Data!A2`, not `Data!A2:A2`), as LibreOffice saves it, in formulas and in defined names.
- A reference with an end that went (`A1:#REF!`, `Data!A1:#REF!`) is an error of that end in the formula of an XLSX and an ODS (`COUNT` 0, `COUNTA` 1, `SUM` the unknown name of an XLSX or the `#REF!` of an ODS), as LibreOffice reads the file it saved.
- The copies of a shared group of an XLSX that would read outside the sheet are written with a formula of their own and a bare `#REF!` when the run writes the sheet, as LibreOffice saves them; a formula of an XLSX that holds a name LibreOffice cannot resolve (a bare `#REF!`) and caches an error is calculated in the default mode, as LibreOffice does when it opens the file.
- The ODS recalculation on load of formulas whose cached result is a plain number, and `ROW()` of an XLSX, are recorded as approved differences of the default mode; the other classes stay proposed, with their cost.

### Fixed

- The range of an ODS with an end that went (`[$Data.A1:.#REF!]`) was read as a reference to an external workbook.
- `sheetrs-io` did not build without its `write` feature (the initializer of `Book` set `calculate_all` only with it).

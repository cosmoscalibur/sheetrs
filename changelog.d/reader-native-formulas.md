### Changed

- The readers give the formula of a cell in **the syntax of its file**: Excel's in an XLSX,
  `OpenFormula` in an ODS (`[.A1]+[$T.B2]`, `;` between arguments, without the `of:=`).
  `read_workbook_into`, `read_selection_into` and `sheetrs query`'s `formula` property work that way.
  The tools that only analyse formulas (linter, statistics, compiler) work on an *analysis form*,
  Excel's notation for both formats: `Source::read_analysis_into` (with the new sink adapter
  `ExcelFormulas`, and `CellSink::sheet`, called when a sheet begins) and `Source::read_workbook`,
  whose `Workbook` is that form, so their results do not change.
- `sheetrs-query` and `sheetrs-edit` see a formula as `LibreOffice` shows it *for that file*
  (`sheetrs_formula::calc_read::native_to_calc`): an ODS keeps sheets that are relative
  (`T.A1`) where an XLSX has only absolute ones (`$T.A1`, also for the own sheet), an ODS range that
  names its sheet twice stays so (`T.A1:T.B3`), the spacing of the file is kept, a function the
  application does not have is shown as written (`foo()`) and the namespace of one it has is left out
  (`COM.MICROSOFT.CONCAT` is `CONCAT`). A replacement writes the result back in the syntax of the
  file.
- `sheetrs-edit`: a replacement that leaves text that is not a formula is made as `LibreOffice`
  makes it (`sheetrs_formula::calc_read::compile_calc`): the cell keeps the text as a formula in an
  error cell, the parentheses left open are closed, the names it does not know are written in
  lowercase and make `#NAME?`, and the error is `Err:508` to `Err:511` or `Err:520` by what is
  wrong. An XLSX writes the errors as Excel's, an ODS as `Err:5xx`. Probed in 26.8 and checked in both
  formats by fixtures made by `LibreOffice` (`tests/assets/replace/make_fixtures.py`). The report no
  longer lists "not a formula" among the cells it skips.

### Fixed

- `xlsx`: a range in a cell that holds `#VALUE!` (`=A1:A3`) reads as `#VALUE!`, as `LibreOffice` shows
  it. The reader took any cached `#VALUE!` of a formula with a `:` for the placeholder of an array
  formula and gave an empty cell; now only an array formula (`t="array"`, which is also how dynamic
  arrays are written) has that placeholder.

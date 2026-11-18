# ODS formulas

Formula cells read by `sheetrs-ods`. The reader keeps the text of `table:formula` as it is, with
its namespace prefix: the grammar is the one the prefix names (`of:` is OpenFormula).

## Observations

| Behaviour | Test |
| --- | --- |
| Of the cached results of formula cells only a `string` with a non-empty `office:string-value` is kept; numbers, booleans, dates, errors and strings without that attribute are recalculated. | `formula_cells_over_ods` in `sheetrs-ods` |
| A text that is not a formula, a `table:formula` with an unknown prefix included, makes the cell an error; the text is kept as the formula. | `formula_cells_over_ods` in `sheetrs-ods` |
| An array formula is a formula cell with both `table:number-matrix-columns-spanned` and `table:number-matrix-rows-spanned`, each an integer after the rules of repeat counts. It fills the area they give, from the cell, with its text, as the area of an XLSX array formula is filled, and an area that reaches past the edge of the sheet is cut there. A size that is missing in one of them, a column count over 32767 and a count that is not positive drop the cell. | `matrix_formulas_fill_their_area_over_ods` and `repeats_and_matrices_at_the_edge_of_the_sheet_over_ods` in `sheetrs-ods` |

## Behaviour of the model

A formula cell is flagged `stale` when LibreOffice recalculates it; its value is the cached one
until the formula engine exists.

## Open defects

- A formula that does not compile (`of:=1+`), or that is volatile (`of:=NOW()`), is
  recalculated by LibreOffice even with a cached string; the reader keeps the string.

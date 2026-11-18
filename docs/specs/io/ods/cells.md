# ODS cell placement

Where `sheetrs-ods` places the cells of a table.

## Observations

| Behaviour | Test |
| --- | --- |
| `number-columns-repeated` and `number-rows-repeated` are the integer their text starts with, after XML whitespace and a sign; a value that is not positive, text without a number and a number beyond 32 bits count as one. A cell with content repeats across its columns and a row repeats down its rows, up to the edge of the sheet. | `rows_columns_and_repeats_over_ods` in `sheetrs-ods` |
| A cell with a formula stands for one column, whatever its repeat count says. | `formula_cells_over_ods` in `sheetrs-ods` |
| A covered cell holds what it says like any other cell; a cell with only a style advances the column. | `rows_columns_and_repeats_over_ods` in `sheetrs-ods` |
| `table-row-group`, `table-header-rows` and `table-rows`, nested in any way, hold rows as the table does. | `rows_columns_and_repeats_over_ods` in `sheetrs-ods` |
| A row inside a row, a cell inside a cell, columns and shapes add no cells and no rows. | `rows_columns_and_repeats_over_ods` in `sheetrs-ods` |
| A repeated cell or row that reaches the edge of the sheet is cut there. | `repeats_and_matrices_at_the_edge_of_the_sheet_over_ods` in `sheetrs-ods` |

## Open defects

- The reader refuses a table that holds more than 67108864 cells, repeated and array cells
  included, because repeat counts make a small file stand for far more cells than it has bytes;
  LibreOffice loads them, as far as memory allows.

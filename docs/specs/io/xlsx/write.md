# XLSX writing

How `LibreOffice` 26.8 writes a workbook.

## Observations

| Behaviour | Test |
| --- | --- |
| XLSX export writes formula results with 15 significant digits: `=0.1+0.2` is stored as `0.3`. | `a_saved_workbook_reads_back_over_xlsx` in `sheetrs-xlsx` |
| A sheet that is very hidden is written as hidden, by `LibreOffice` as well, in XLSX and in ODS. | `very_hidden_sheets_are_saved_as_hidden_over_xlsx` in `sheetrs-xlsx` |
| Only the first cell of an array formula is written as a formula, with the `ref` of its area; the other cells hold no formula of their own. | `only_the_first_cell_of_an_array_is_written_as_a_formula_over_xlsx` in `sheetrs-xlsx` |
| Numbers are stored with 15 significant digits taken from the shortest decimal that reads back as the number, rounded half up at the sixteenth digit; in full from `1E-4` up to `1E+17` (the 16 integer digits of a number from `1E+15` are not rounded), and as `1.5E+017` or `1E-005` outside. | `stored_numbers_are_those_of_the_oracle_over_xlsx` in `sheetrs-xlsx` |
| `LibreOffice` loads what the writer produces as the reader reads it: numbers, text with escapes and control characters, booleans, errors and formulas. | `written_workbooks_load_as_they_read_over_xlsx` in `sheetrs-xlsx` |

What the writer keeps from the package a workbook was read from is in [opaque content](../../model/opaque.md).

## Open defects

| Defect | Closes in | Why it waits |
| --- | --- | --- |
| Formula text is written as the model holds it: the normalization `LibreOffice` applies when it saves (spacing, case of names, references) is not reproduced. | M2 | Normalization needs the formula parser that M2 builds; reproducing it from text patterns would be a second parser. |
| Number formats and styles of workbooks that were not read from an XLSX package are not written: the minimal package has one default style. | Not scheduled: needs a decision. | The model holds no styles; they only travel as opaque content. |
| A formula that `LibreOffice` calculates again when it loads is written with the result the workbook holds, which is the cached one: `LibreOffice` writes the one it calculated, which differs from the cached one in the last digit in a small share of the cells of the private sets, and so do the cells that depend on them, directly or not, which `LibreOffice` calculates again too. | M2 | The result is what the formula engine of M2 calculates. |

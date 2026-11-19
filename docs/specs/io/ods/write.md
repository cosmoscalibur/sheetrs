# ODS writing

How `LibreOffice` 26.8 writes a workbook as ODS.

## Observations

| Behaviour | Test |
| --- | --- |
| `office:value` holds a number with 15 significant digits, taken from the shortest decimal that reads back as the number and rounded half up at the sixteenth. It is written in full from `1E-14` up to `1E+15` and below 2^53, and with an exponent of three digits outside (`1E+016`, `9E-016`); a negative zero is `-0`. | `stored_numbers_are_those_of_the_oracle_over_ods` in `sheetrs-ods` |
| A workbook read from an ODS package and not changed is written with every entry as it was read, `content.xml` included, byte for byte. | `an_unchanged_workbook_is_written_as_it_was_read_over_ods` in `sheetrs-ods` |
| A changed workbook keeps every entry but `content.xml`, which is written from the model. | `a_changed_workbook_keeps_the_other_parts_over_ods` in `sheetrs-ods` |
| `LibreOffice` loads what the writer produces as the reader reads it: numbers, text with runs of spaces, tabs and line breaks, booleans, errors, formulas and the gaps between cells and rows. | `written_workbooks_load_as_they_read_over_ods` in `sheetrs-ods` |
| Names, hidden tables and the date system of a written package load as the reader reads them. | `written_names_tables_and_dates_load_as_they_read_over_ods` in `sheetrs-ods` |
| A matrix is written as one formula in its first cell, with the size of its area, and the cells it covers as `table:covered-table-cell`. `LibreOffice` drops the characters XML 1.0 does not allow from text, and writes `#VALUE!` for errors it has no name for. | `written_matrices_and_odd_text_load_as_they_read_over_ods` and `matrices_and_odd_text_are_written_as_libreoffice_does_over_ods` in `sheetrs-ods` |
| A constant error is not a value of an ODS cell: it is the formula that is that error, and the reader leaves its cached text to be calculated again. A carriage return in text is read as a line feed. | `written_workbooks_read_back_over_ods` in `sheetrs-ods` |

`LibreOffice` writes a sheet that is very hidden in the source as hidden (see [XLSX writing](../xlsx/write.md)), which is what the writer does.

## Open defects

| Defect | Closes in | Why it waits |
| --- | --- | --- |
| A workbook whose formulas are in another grammar than OpenFormula is refused with `Error::Unsupported`. | M2 | Translating the grammar of a formula needs the formula parser that M2 builds. |
| The styles, widths and merged cells inside `content.xml` are lost when the workbook was changed after it was read: the writer then writes `content.xml` from the model. | M3 | Carrying them through a change means moving them with the cells, which is what the operations of M3 (delete rows and columns, replace formulas by values) define for every format; designing it before them would fix a behaviour the oracle has not shown yet. |
| A formula that `LibreOffice` calculates again when it loads is written with the result the workbook holds, which is the cached one: `LibreOffice` writes the one it calculated, which differs from the cached one in the last digit in a small share of the cells of the private sets, and so do the cells that depend on them, directly or not, which `LibreOffice` calculates again too. | M2 | The result is what the formula engine of M2 calculates. |

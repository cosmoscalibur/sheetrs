# ODS writing

How `LibreOffice` 26.8 writes a workbook as ODS.

## Observations

| Behaviour | Test |
| --- | --- |
| `office:value` holds a number with 15 significant digits, taken from the shortest decimal that reads back as the number and rounded half up at the sixteenth. It is written in full from `1E-14` up to `1E+15` and below 2^53, and with an exponent of three digits outside (`1E+016`, `9E-016`); a negative zero is `-0`. | `stored_numbers_are_those_of_the_oracle_over_ods` in `sheetrs-ods` |
| `LibreOffice` loads what the writer produces as the reader reads it: numbers, text with runs of spaces, tabs and line breaks, booleans, errors, formulas and the gaps between cells and rows. | `written_workbooks_load_as_they_read_over_ods` in `sheetrs-ods` |
| Names, hidden tables and the date system of a written package load as the reader reads them. | `written_names_tables_and_dates_load_as_they_read_over_ods` in `sheetrs-ods` |
| A constant error is not a value of an ODS cell: it is the formula that is that error, and the reader leaves its cached text to be calculated again. A carriage return in text is read as a line feed. | `written_workbooks_read_back_over_ods` in `sheetrs-ods` |

## Open defects

| Defect | Closes in | Why it waits |
| --- | --- | --- |
| A workbook whose formulas are in another grammar than OpenFormula is refused with `Error::Unsupported`. | M2 | Translating the grammar of a formula needs the formula parser that M2 builds. |
| A very hidden sheet is written as hidden. | M1 | The oracle has to show what `LibreOffice` writes for it. |
| What the package carries beyond the model (styles, widths, the other parts) is not written yet. | M1 | It is the next step of the writer: the same carry-over as XLSX, for ODS. |

# CSV writing

How `LibreOffice` 26.8 exports a sheet with the text filter, with the options of its dialog left
as they come: a comma, `"` as text delimiter, UTF-8, cells as shown.

## Observations

| Behaviour | Test |
| --- | --- |
| A line per row, from the first row to the last one that has a cell, ended by `\n`; every line has a field for each column up to the last one that has a cell. | `exported_text_is_what_the_oracle_exports_over_csv` in `sheetrs-csv` |
| Text is quoted only when it holds the separator, the delimiter or a line break, and a delimiter inside it is doubled. | same |
| Numbers show as General format shows them: 15 significant digits, from `1E-4` to `1E+15` in full, in full from `1E-9` when the decimals do not pass 16, and as `1.5E+016` or `1E-10` otherwise. The first digit of a number below `1E+16` that has 16 integer digits is not rounded. | same |
| The 15 digits come from the shortest decimal that reads back as the number, rounded half up at the sixteenth digit: `3395.843825702095` shows as `3395.8438257021`. | same |
| Booleans are `TRUE` and `FALSE` and errors their text, `Err:502` among them. | same |

## Open defects

No open defects for the options above. The export options that the dialog offers beyond the
separator and the delimiter (quoting every text cell, saving formulas, a byte order mark, other
character sets, a sheet other than the first) are not written.

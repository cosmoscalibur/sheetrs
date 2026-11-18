# XLSX cell values

How `LibreOffice` 26.8 reads the value of a `<c>` element (`sheetrs-xlsx`).

## Observations

| Behaviour | Test |
| --- | --- |
| `_xHHHH_` escapes are decoded in shared strings and inline strings (`_x000D_` is a carriage return, `_x005F_` an underscore, two surrogate escapes in a row one character, a lone surrogate escape ends the text before it) but not in the result of a `t="str"` cell. | `typed_cell_values_over_xlsx` in `sheetrs-xlsx` |
| In shared and inline strings a carriage return and a line feed in either order are one line feed; a result of `t="str"` keeps them. | `typed_cell_values_over_xlsx` in `sheetrs-xlsx` |
| An empty `<v>` or `<v/>` holds nothing, for any cell type; an empty shared string reads as an empty cell. | `typed_cell_values_over_xlsx` in `sheetrs-xlsx` |
| Inline text counts only in `inlineStr` cells; such a cell with an empty `<is>`, without `<is>` or with a `<v>` element holds nothing. | `typed_cell_values_over_xlsx` in `sheetrs-xlsx` |
| A stored number is its longest decimal prefix after leading spaces and tabs; text without one is `0`, keeping the sign (`-` is `-0`). `NaN` (unsigned), `1.#NAN`, `INF` and `1.#INF` are not-a-number and infinity, and `1e999` is infinity. | `typed_cell_values_over_xlsx` in `sheetrs-xlsx` |
| A boolean cell is `TRUE` for the word `true` in any case or a number other than zero (`NaN` and infinity included), and `FALSE` otherwise. | `typed_cell_values_over_xlsx` in `sheetrs-xlsx` |
| A shared string reference is an integer prefix after whitespace (space, tab, CR, LF, U+2000 to U+200B, U+2028, U+2029); text without one, or a number beyond 32 bits, is `0`; a negative number or an index past the table is an empty cell. | `typed_cell_values_over_xlsx` in `sheetrs-xlsx` |
| A cell with an unknown type holds nothing. | `typed_cell_values_over_xlsx` in `sheetrs-xlsx` |
| A cell of type `d` is a serial number counted from 1899-12-30 (1904-01-01 in the 1904 system) when its trimmed text is `YYYY-MM-DD`, optionally followed by `T` or `t`, `hh:mm:ss`, a fraction after `.` or `,` and `Z` or an offset. The year has one or more digits, an optional minus sign and a value from -32768 to 32767 other than 0; fractional seconds are dropped; hour 24 needs zero minutes, seconds and leading fraction; a positive offset counts its hours only, a negative one counts its minutes twice, up to 14:00. Other text stays text. | `typed_cell_values_over_xlsx` in `sheetrs-xlsx` |

| A boolean cell is the formula `=TRUE()` or `=FALSE()`, and an error cell whose text names an error (the six classic errors in any case, ignoring ASCII whitespace, and `#SPILL!`) is the formula of that name; both are recalculated on load. | `typed_cell_values_over_xlsx` and `cell_addresses_rows_and_duplicates_over_xlsx` in `sheetrs-xlsx` |

## Open defects

- A character that XML 1.0 forbids, such as U+000B, inside a cell value makes LibreOffice
  read the cell as empty; the reader keeps the text.
- An error cell without a formula is compiled by LibreOffice as a formula whose text is
  decompiled in its own way (`1e999` is `=INF`, `a1` is `=A1`, `#ref!` is a name that gives
  `#NAME?`). The reader gives every other text the error LibreOffice shows, with no formula;
  text such as `1`, `TRUE`, `"a"`, `A1`, `(` or `=` gives other values. Closing it needs the
  formula compiler.
- A boolean or error cell at the address of a cell that an array formula filled replaces it in
  the reader and leaves it in LibreOffice.

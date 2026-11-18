# ODS cell values

How `LibreOffice` 26.8 reads the value of a `table:table-cell` (`sheetrs-ods`).

## Observations

| Behaviour | Test |
| --- | --- |
| The type of a cell is its `calcext:value-type` when it has one and it is not `error`, and its `office:value-type` otherwise. A type that is not `float`, `percentage`, `currency`, `string`, `boolean`, `date` or `time`, an empty `calcext:value-type` included, makes the cell empty, whatever its attributes and text. A cell without a type but with paragraphs is a string. | `typed_cell_values_over_ods` in `sheetrs-ods` |
| `float`, `percentage` and `currency` cells hold the number their `office:value` starts with, after optional XML whitespace and a sign: the longest decimal prefix, `-0` being `-0`. Text without such a prefix, the words `inf` and `NaN` and a number that overflows are `0`; a cell without `office:value` is `0`, whatever its paragraphs say. | `typed_cell_values_over_ods` in `sheetrs-ods` |
| A `boolean` cell is `1` for the `office:boolean-value` text `true` (lower case only), and otherwise the number that text starts with (so `2` is `2` and `TRUE` is `0`); without that attribute it reads `office:value`, and without either it is `0`. LibreOffice holds a number; the model holds a boolean for 0 and 1 and a number for any other value. | `typed_cell_values_over_ods` in `sheetrs-ods` |
| A `date` cell is the serial number of its `office:date-value`: `YYYY-MM-DD`, optionally followed by `T` or `t`, `hh:mm:ss`, a fraction after `.` or `,` and `Z` or an offset, between whitespace. The year has one or more digits, an optional minus sign and a value from -32768 to 32767 other than 0, and a year before 1 counts as in the calendar without a year 0. The seconds keep their fraction; hour 24 needs zero minutes, seconds and leading fraction; a positive offset counts its hours only, a negative one counts its minutes twice, up to 14:00. Text that is not a date is `0`; without the attribute the cell reads `office:value`. | `typed_cell_values_over_ods` in `sheetrs-ods` |
| A `time` cell is the days of its `office:time-value`, an ISO 8601 duration `[-]P[nD][T[nH][nM][nS]]` in either case, with the seconds as a fraction of the day of their total. Years and months, a sign inside the text, spaces, and an amount of 2147483640 or more make it `0`. A repeated unit counts once, the last time. A decimal point after any amount of the time part makes the amount seconds, whatever its unit; after the days it makes the duration `0`. Without the attribute the cell reads `office:value`. | `typed_cell_values_over_ods` in `sheetrs-ods` |
| A `string` cell is its non-empty `office:string-value`; otherwise its paragraphs joined by line feeds, where one empty paragraph is an empty string and no paragraph is an empty cell. | `text_of_string_cells_over_ods` in `sheetrs-ods` |
| Text is read from the paragraphs (`text:p`) that are children of the cell; headings, lists and annotations add nothing. `text:s` is its `text:c` spaces (one when the count is missing, zero, negative or not a number), `text:tab` a tab and `text:line-break` a line feed; a carriage return, with or without a line feed, is a line feed. Elements that are not known are read as if their content were in the paragraph. | `text_of_string_cells_over_ods` in `sheetrs-ods` |
| In a `text:span` only text and `text:s` count, not tabs and line breaks; a span inside a span, or inside a link, adds nothing; a link has the same limits and takes one space when it has no text; `text:ruby-text` adds nothing. | `text_of_string_cells_over_ods` in `sheetrs-ods` |
| Elements and attributes are told apart by namespace, not by prefix. | `namespaces_are_resolved_by_uri_over_ods` in `sheetrs-ods` |

## Open defects

- The reader refuses a paragraph whose text grows past 16777216 bytes through `text:s` and
  elements nested more than 256 deep inside a paragraph; the limits of LibreOffice are not
  known.
- Fields of text inside a paragraph (`text:date` and the like) are read as their content, while
  `LibreOffice` evaluates them.

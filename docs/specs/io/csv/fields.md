# Delimited text fields

How `LibreOffice` 26.8 imports a delimited text file with the text filter, whose options each
case sets explicitly (`sheetrs-csv`). The filter reads UTF-8, English (United States), with the
detection of special numbers off.

## Observations

| Behaviour | Test |
| --- | --- |
| A field is a number when it is `[+-]digits[.digits][e[+-]digits]` or `[+-].digits[...]` between ASCII spaces and no-break spaces, whatever its leading zeros; `1.` and `.5` are numbers, `.e5` and `1e` are not. A number that overflows or underflows to zero is text. Other fields are text, the words `NaN`, `inf`, `TRUE` and `#N/A` and numbers in other digits included, and so are fields with other whitespace around a number. A field that is only spaces is text. | `plain_numbers_and_text_over_csv` in `sheetrs-csv` |
| The text of a field is kept as it is, with its spaces; an empty field holds nothing. | `quotes_line_breaks_and_rows_over_csv` in `sheetrs-csv` |
| A quote at the start of a field, after spaces, opens a quoted field. It closes at the first quote that is followed, after spaces, by a separator, a line break or the end of the text; two quotes in a row are one, and any other quote is content. A field with no such quote is not quoted: its text, quotes included, runs to the next separator or line break. Spaces between the closing quote and what follows are part of the text. | `quotes_line_breaks_and_rows_over_csv` in `sheetrs-csv` |
| A carriage return, a line feed, a carriage return followed by a line feed and a line feed followed by a carriage return are each one line break; in a quoted field each is a line feed. A byte order mark at the start is dropped, and so are the null characters; bytes that are not UTF-8 are the replacement character (`bytes_that_are_not_utf8_over_csv`). | `quotes_line_breaks_and_rows_over_csv` in `sheetrs-csv` |
| A row can be shorter or longer than the one before it. Columns and rows beyond the sheet are dropped. | `quotes_line_breaks_and_rows_over_csv` and `the_edge_of_the_sheet_over_csv` in `sheetrs-csv` |
| With quoted fields as text, a quoted field is text even when it looks like a number. | `quoted_fields_as_text_over_csv` in `sheetrs-csv` |
| With spaces trimmed, the ASCII spaces around a field that is not quoted are dropped, and so are the ones after the closing quote; other whitespace stays. | `trimmed_spaces_over_csv` in `sheetrs-csv` |
| With formulas evaluated, a field of more than one character that starts with `=`, quoted or not, is a formula to recalculate; the text stays as the field has it. | `formula_fields_over_csv` in `sheetrs-csv` |
| The separator and the delimiter are any single characters, the delimiter can be none, and a space as separator is never one of the spaces around a quoted field. | `other_separators_and_delimiters_over_csv` in `sheetrs-csv` |

## Open defects

- The detection of special numbers is not implemented: with it on, `LibreOffice` also reads
  dates and times in the formats of its language, percentages, one currency, booleans, and
  numbers with a group separator as numbers; some of them depend on the current year.
- A separator that is the delimiter, a separator that is a letter and a delimiter that is a
  space are not read as LibreOffice reads them; the tests cover the separators `,` `;` tab `|`
  and space, and the delimiters `"`, `'` and none.
- Only UTF-8 is read.
- The whole input is held in memory while it is read.

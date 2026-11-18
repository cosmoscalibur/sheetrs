# Oracle observations

Behaviour of LibreOffice 26.8 that is not stated in a public specification,
with the oracle test that shows it.

| Behaviour | Test |
| --- | --- |
| Formula to Value turns a formula that evaluates to an error into an empty text cell. | `converting_an_error_formula_to_a_value_leaves_an_empty_string` in `sheetrs-testkit` |
| Formula to Value converts the cached results, so the cells must be calculated first. | `oracle.py`, operation `to_values` |
| XLSX export writes formula results with 15 significant digits: `=0.1+0.2` is stored as `0.3`. | `a_workbook_saved_by_the_oracle_reads_back_to_the_observed_values` in `sheetrs-xlsx` |
| Of the cached results of formula cells only a plain formula's non-zero number is trusted on load; strings, booleans, errors, dates, zero, shared and array formulas, and missing caches are recalculated, and a data table is an empty cell. | `only_the_cache_of_a_plain_non_zero_number_is_trusted_on_load` in `sheetrs-xlsx` |
| `_xHHHH_` escapes are decoded in shared strings and inline strings (`_x000D_` is a carriage return, `_x005F_` an underscore, two surrogate escapes in a row one character, a lone surrogate escape ends the text before it) but not in the result of a `t="str"` cell. | `crafted_cells_read_as_the_oracle_loads_them` in `sheetrs-xlsx` |
| In shared and inline strings a carriage return and a line feed in either order are one line feed; a result of `t="str"` keeps them. | `crafted_cells_read_as_the_oracle_loads_them` |
| An empty `<v>` or `<v/>` holds nothing, for any cell type; an empty shared string reads as an empty cell. | `crafted_cells_read_as_the_oracle_loads_them` |
| Inline text counts only in `inlineStr` cells; such a cell with an empty `<is>`, without `<is>` or with a `<v>` element holds nothing. | `crafted_cells_read_as_the_oracle_loads_them` |
| A stored number is its longest decimal prefix after leading spaces and tabs; text without one is `0`, keeping the sign (`-` is `-0`). `NaN` (unsigned), `1.#NAN`, `INF` and `1.#INF` are not-a-number and infinity, and `1e999` is infinity. | `crafted_cells_read_as_the_oracle_loads_them` |
| A boolean cell is `TRUE` for the word `true` in any case or a number other than zero (`NaN` and infinity included), and `FALSE` otherwise. | `crafted_cells_read_as_the_oracle_loads_them` |
| A shared string reference is an integer prefix after whitespace (space, tab, CR, LF, U+2000 to U+200B, U+2028, U+2029); text without one, or a number beyond 32 bits, is `0`; a negative number or an index past the table is an empty cell. | `crafted_cells_read_as_the_oracle_loads_them` |
| A cell with an unknown type holds nothing. | `crafted_cells_read_as_the_oracle_loads_them` |
| A cell of type `d` is a serial number counted from 1899-12-30 (1904-01-01 in the 1904 system) when its trimmed text is `YYYY-MM-DD`, optionally followed by `T` or `t`, `hh:mm:ss`, a fraction after `.` or `,` and `Z` or an offset. The year has one or more digits, an optional minus sign and a value from -32768 to 32767 other than 0; fractional seconds are dropped; hour 24 needs zero minutes, seconds and leading fraction; a positive offset counts its hours only, a negative one counts its minutes twice, up to 14:00. Other text stays text. | `crafted_cells_read_as_the_oracle_loads_them` |
| The `date1904` attribute is on for the words `true`, `on` and `t`, or a number other than zero after optional whitespace; the last workbook properties element wins. | `crafted_cells_read_as_the_oracle_loads_them` |
| A cell address that is not a cell of the sheet (`XFE1`, `A0`, `1A`, `B1:C1`, `Sheet1!B1`, empty) places the cell like a cell without an address, at the next column of the current row. | `crafted_cells_read_as_the_oracle_loads_them` |
| A row attribute is an integer prefix after XML whitespace (space, tab, CR, LF); text without digits and `0` make a row outside the sheet whose unaddressed cells are dropped, and the next row follows it; `-1` is a row without the attribute; addressed cells ignore the row. | `crafted_cells_read_as_the_oracle_loads_them` |
| The `t` of a formula is `normal`, `shared`, `array` or `dataTable`, written exactly; any other value, and a shared group below zero, drops the formula and keeps the cell value as a constant. The group number is an integer prefix and a shared formula without a group is an ordinary one. | `crafted_cells_read_as_the_oracle_loads_them` |
| An array formula fills every cell of its `ref` with its text, as cells to recalculate: later cells that hold a formula replace it, later constants do not, and earlier cells in the area are replaced. The `ref` has to start at the cell that holds the formula, in the grammar of the shared `ref`, or the cell is dropped with its cached value; a range whose end is a column alone or an incomplete row fills only its first row, and a reversed range counts as its normalized form. An empty text fills the area too. Areas are cut at the sheet. | `crafted_cells_read_as_the_oracle_loads_them` |
| A shared formula group is kept only when one of its anchors (a cell with formula text or a `ref`) has a `ref` that covers the anchor itself; otherwise the anchor and every dependent of the group are empty cells, a dependent whose group has no anchor included. One valid anchor keeps the whole group, whichever order the anchors come in. A member of a dropped group writes nothing: where it shared an address with another cell, in either order, that cell stays. | `crafted_cells_read_as_the_oracle_loads_them` |
| A `ref` is a cell, or two corners (cells, columns or rows) in either order, followed by spaces only (a tab, line feed or no-break space, written as a character reference, rejects it); letters in any case, `$` before a column or a row, a `+` before a row, leading zeros, and rows beyond the sheet are accepted in the end corner. The end corner is lenient: with a column valid in the sheet followed by a digit or a sign and a row that is zero, negative or beyond the sheet, anything after the column is ignored and the range runs to the last row; a row inside the sheet must be followed by nothing. The end corner may be a column alone, with or without `$`; a column alone or a row alone is not a cell reference on its own. | `crafted_cells_read_as_the_oracle_loads_them` |
| When two cells share an address, a formula cell is not replaced by a later cell without formula; otherwise the later cell wins. | `crafted_cells_read_as_the_oracle_loads_them` |
| A worksheet cut short keeps the cells read before the cut. | `a_worksheet_cut_at_a_tag_boundary_keeps_the_cells_read` in `sheetrs-xlsx` |

Every difference between the reader and the oracle is a defect.

## Open defects

Differences that are known and not fixed yet, with the inputs that show them:

- A boolean and an error cell without a formula are formula cells in LibreOffice
  (`=TRUE()`, `=FALSE()`, `=#N/A`); the reader returns constants with the same value.
- An error cell without a formula is compiled by LibreOffice as a formula. The reader maps
  the six classic errors (ignoring case and ASCII whitespace), `#SPILL!`, `#GETTING_DATA`
  and a classic error followed by other text, and reads everything else as `#NAME?`; text
  such as `1`, `TRUE`, `"a"`, `A1`, `(`, `=` or `#N/A!` gives other values or error codes.
- An array formula over the whole sheet (`A1:XFD1048576`) makes LibreOffice fail to load the file; the reader refuses an area of more than `MAX_ARRAY_CELLS` cells. LibreOffice loads areas of 27 million cells, whose limit depends on the memory of the machine.
- A plain formula with a non-zero number cached is flagged as trusted, but LibreOffice
  recalculates it when its text does not compile (`<f>  </f>` is `Err:520`), when it is volatile
  (`NOW()`) and when it is circular or evaluates to an error.
- LibreOffice reads element and attribute namespaces and element nesting: it ignores `x:r`,
  `x:t`, elements in a foreign namespace, a worksheet whose root is not in the main namespace,
  a `<c>` outside a `<row>`, a `<row>` outside `<sheetData>` and nested rows or cells, restarts
  the row counter at each `<sheetData>` and keeps the last non-empty `<v>`; the reader matches
  names without namespaces or nesting.
- LibreOffice loads files with undefined entities, character references to forbidden
  characters, an unknown sheet `state` and a sheet whose relationship is missing (as an empty
  sheet); the reader fails on them, and on a workbook part, relationships part or shared
  strings part that is cut short. Worksheets with mismatched or unmatched end tags and repeated
  attributes are read without error, and which cells they keep has not been compared.
- LibreOffice renames sheets with invalid names and loads duplicate names as separate sheets;
  the reader keeps the names as stored.
- A character that XML 1.0 forbids, such as U+000B, inside a cell value makes LibreOffice
  read the cell as empty; the reader keeps the text.
- The day count of dates before year 1 differs from LibreOffice's by one day for some
  years (`-4712-01-01`, `-32768-01-01`, `-2020-01-02`); the reader counts the proleptic
  Gregorian calendar.
- Formula cells that LibreOffice recalculates on load are flagged `stale`; their value is
  the cached one until the formula engine exists.

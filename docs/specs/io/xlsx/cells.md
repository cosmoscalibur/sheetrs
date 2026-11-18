# XLSX cell placement

Where `sheetrs-xlsx` places the cells of a worksheet: addresses, rows, repeated addresses and truncated parts.

## Observations

| Behaviour | Test |
| --- | --- |
| A cell address that is not a cell of the sheet (`XFE1`, `A0`, `1A`, `B1:C1`, `Sheet1!B1`, empty) places the cell like a cell without an address, at the next column of the current row. | `cell_addresses_rows_and_duplicates_over_xlsx` in `sheetrs-xlsx` |
| A row attribute is an integer prefix after XML whitespace (space, tab, CR, LF); text without digits and `0` make a row outside the sheet whose unaddressed cells are dropped, and the next row follows it; `-1` is a row without the attribute; addressed cells ignore the row. | `cell_addresses_rows_and_duplicates_over_xlsx` in `sheetrs-xlsx` |
| When two cells share an address, a formula cell is not replaced by a later cell without formula; otherwise the later cell wins. | `cell_addresses_rows_and_duplicates_over_xlsx` in `sheetrs-xlsx` |
| A worksheet cut short keeps the cells read before the cut. | `a_worksheet_cut_at_a_tag_boundary_keeps_the_cells_read_over_xlsx` in `sheetrs-xlsx` |

## Open defects

- LibreOffice reads element and attribute namespaces and element nesting: it ignores `x:r`,
  `x:t`, elements in a foreign namespace, a worksheet whose root is not in the main namespace,
  a `<c>` outside a `<row>`, a `<row>` outside `<sheetData>` and nested rows or cells, restarts
  the row counter at each `<sheetData>` and keeps the last non-empty `<v>`; the reader matches
  names without namespaces or nesting.
- Worksheets with mismatched or unmatched end tags and repeated attributes are read without
  error by the reader and by LibreOffice; which cells they keep has not been compared.

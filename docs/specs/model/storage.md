# Storage limits

Capacities of the cell store (`sheetrs-core`) that can make a read fail where
`LibreOffice` succeeds, or the reverse.

## Observations

| Behaviour | Test |
| --- | --- |
| An array formula covering the whole sheet (`A1:XFD1048576`) makes LibreOffice fail to load the file. | `an_array_formula_that_covers_too_many_cells_is_refused` in `sheetrs-xlsx` |

## Open defects

- The store refuses an array formula of more than `MAX_ARRAY_CELLS` cells with
  `LimitError::ArrayCells`. LibreOffice loads an area of 27 million cells, and where it stops
  depends on the memory of the machine, so the cap is a choice, not an observation.
- The cap is per array, not per sheet: several overlapping arrays can each fill up to the cap.

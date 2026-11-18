# Cells in the model

What a cell holds, and how the model differs from `LibreOffice` in it.

## Formula text

A formula keeps the text the file stores, so that its grammar is the one of the format: an XLSX
`<f>` element without `=`, the `table:formula` attribute of ODS with its namespace prefix and
`=` (`of:=SUM([.B1];2)`), and a field of a text file with its `=`.

## Observations

| Behaviour | Test |
| --- | --- |

## Open defects

- Formula cells that LibreOffice recalculates on load are flagged `stale`; their value is
  the cached one until the formula engine exists.

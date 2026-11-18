# XLSX writing

How `LibreOffice` 26.8 writes a workbook.

## Observations

| Behaviour | Test |
| --- | --- |
| XLSX export writes formula results with 15 significant digits: `=0.1+0.2` is stored as `0.3`. | `a_saved_workbook_reads_back_over_xlsx` in `sheetrs-xlsx` |

## Open defects

No open defects: the writer does not exist yet.

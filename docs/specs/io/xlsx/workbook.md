# XLSX workbook and sheets

Workbook-level parts read by `sheetrs-xlsx`: properties, sheet list and relationships.

## Observations

| Behaviour | Test |
| --- | --- |
| The `date1904` attribute is on for the words `true`, `on` and `t`, or a number other than zero after optional whitespace; the last workbook properties element wins. | `date_cells_follow_the_date_system_flag_over_xlsx` in `sheetrs-xlsx` |

| A sheet is hidden when its `state` is exactly `hidden` or `veryHidden`; any other state, in any case, is visible. | `sheet_states_over_xlsx` in `sheetrs-xlsx` |
| The sheet the workbook opens on is shown whatever its state says: the one the integer that `activeTab` of the first `workbookView` starts with names, or the first sheet when that is not a sheet. | `sheet_states_over_xlsx` in `sheetrs-xlsx` |

## Open defects

- LibreOffice loads files with undefined entities, character references to forbidden
  characters and a sheet whose relationship is missing (as an empty sheet); the reader fails on them, and on a workbook part, relationships part or shared
  strings part that is cut short.
- LibreOffice renames sheets with invalid names and loads duplicate names as separate sheets;
  the reader keeps the names as stored.

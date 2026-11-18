# ODS workbook and tables

## Observations

| Behaviour | Test |
| --- | --- |
| Serial numbers count days from the `table:date-value` of `table:null-date`, any valid date; text that is not a date leaves the zero at 1899-12-30. | `dates_follow_the_null_date_over_ods` in `sheetrs-ods` |
| A table is hidden when the first table style of its name has `table:display` exactly `false` in its table properties; any other value, a missing style and a style of another family leave it visible. No table is shown for being the one the document opens on. | `hidden_tables_over_ods` in `sheetrs-ods` |
| A document that is cut short, has end tags that do not match, has content after the root element or an undefined entity cannot be loaded; whitespace after the root is fine. | `documents_that_are_not_well_formed_over_ods` in `sheetrs-ods` |

## Open defects

- Table names are kept as stored, without the renaming that `LibreOffice` applies to names it
  cannot use.

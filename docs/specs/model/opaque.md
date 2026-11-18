# Opaque content

What a workbook carries from its package without interpreting it, so that the writer of the same
format puts it back. A workbook read from an XLSX package holds the package entries the model does
not interpret (styles, drawings, themes, macro projects, external links, document properties), the
text of the workbook part around its list of sheets, and for each sheet the text of its part around
the cells, the attributes of its rows and the style of each cell, including the cells that hold
only a style.

## Observations

| Behaviour | Test |
| --- | --- |
| A package written after being read keeps its unread entries byte for byte, the workbook part around `sheets` and the sheet parts around `sheetData`, the attributes of the rows except `r` and `spans`, and the style of every cell, including a cell without a value. | `what_the_model_does_not_hold_is_written_back_over_xlsx` in `sheetrs-xlsx` |
| `LibreOffice` loads a package that was read and written back with the same cells as the original. | `a_package_written_back_loads_as_it_reads_over_xlsx` in `sheetrs-xlsx` |

## Open defects

An object the model does not interpret travels as opaque content and survives only when the
workbook is written in the format it was read from; written in another format it is lost.
Interpreting such objects in a neutral form, one kind at a time, is roadmap work after M7 (a
version after 0.5.1): it changes no other part of the design, because opaque content sits beside
the cells and not in them. A kind of object moves out of the opaque content earlier when a
workbook of the private sets needs it for correct behaviour.

| Defect | Closes in | Why it waits |
| --- | --- | --- |
| A workbook written in a format other than the one it was read from loses its opaque content. | Roadmap after M7 | See above. |
| Row attributes and cell styles are kept as XLSX text and ids, so they are lost on conversion although Calc models them. | Roadmap after M7 | Their neutral form is the first kind to interpret; it needs the stage that defines the style model. |
| Rich text runs of shared strings are read as plain text. | To be measured against the private sets | Closed in M1 if those workbooks hold rich text. |
| The cell attributes `cm`, `vm`, `ph` and the row attribute `spans` are not written back. | To be measured against the private sets | Closed in M1 if those workbooks hold them. |
| The calculation chain is copied as it was read and no longer matches a workbook whose cells changed. | M3 | The operations that change cells decide whether to drop or rebuild it. |
| Opaque references to cells (defined names, merged cells, conditional formats, data validations, hyperlinks, drawing anchors) are not adjusted when rows, columns or sheets are deleted. | M3 | The adjustment is the semantics of those operations. |
| A sheet without a part in the package, and a date system that differs from the package's, are refused with `Error::Unsupported`. | M3 | Adding a sheet to a package needs its relationships and content types, which the sheet operations define. |

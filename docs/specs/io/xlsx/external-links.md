# XLSX external links

## Observations

| Behaviour | Test |
| --- | --- |
| A workbook refers to the documents that the `externalLink` parts listed by its `externalReferences` point at: the target of the relationship of type `externalLinkPath` that the `externalBook` names. An empty target, a relationship of another type, a missing relationship, an `oleLink` and a `ddeLink` name no document, and neither does a part that `externalReferences` does not list. | `external_links_over_xlsx` in `sheetrs-xlsx` |
| A document that two links name counts once. | `external_links_over_xlsx` in `sheetrs-xlsx` |
| A link that cannot be read does not stop the workbook from loading: a relationships part that is not well-formed names no document, and a link part that goes wrong after its `externalBook` element still names its document. | `external_links_over_xlsx` in `sheetrs-xlsx` |

## Open defects

- A formula that refers to an external book by number (`[1]Sheet1!A1`) makes LibreOffice name
  that book by its number when the workbook does not list it; the reader lists no document, as
  it does not read formulas.
- The reader lists the target as written. LibreOffice names a path relative to the workbook by
  its address in the folder of the workbook, a Windows path as a `file` address and a space as
  `%20`; the test compares the names after those rewrites.
- A link part with an unfinished tag (`<externalBook r:id="rId1"</externalLink>`) names no
  document for LibreOffice, and does for the reader, whose XML parser accepts the tag.

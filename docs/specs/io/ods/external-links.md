# ODS external links

## Observations

| Behaviour | Test |
| --- | --- |
| A document is referred to by a formula that holds a quoted address followed by `#` (`['file:///tmp/a.ods'#$Sheet1.A1]`), in the grammars the tests use (`of:`, `oooc:` and none) and outside the text of the formula; the address is read with its doubled quotes as one. A quoted name without `#`, an address inside a text and a link of the table (`table:table-source`) refer to none. A document that several formulas name counts once. | `external_links_over_ods` in `sheetrs-ods` |

## Open defects

- A formula with such a reference is recalculated by LibreOffice, which gives `Err:540` while the
  document cannot be opened; the reader keeps the cached value, flagged stale.

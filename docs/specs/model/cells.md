# Cells in the model

What a cell holds, and how the model differs from `LibreOffice` in it.

## Observations

| Behaviour | Test |
| --- | --- |

## Open defects

- A boolean and an error cell without a formula are formula cells in LibreOffice
  (`=TRUE()`, `=FALSE()`, `=#N/A`); the model holds constants with the same value.
- An error cell without a formula is compiled by LibreOffice as a formula. The reader maps
  the six classic errors (ignoring case and ASCII whitespace), `#SPILL!`, `#GETTING_DATA`
  and a classic error followed by other text, and reads everything else as `#NAME?`; text
  such as `1`, `TRUE`, `"a"`, `A1`, `(`, `=` or `#N/A!` gives other values or error codes.
- Formula cells that LibreOffice recalculates on load are flagged `stale`; their value is
  the cached one until the formula engine exists.

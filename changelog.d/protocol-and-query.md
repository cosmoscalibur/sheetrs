### Added

- New crate `sheetrs-protocol`: the contract of the tools as typed JSON (`Request`, `Response`,
  `ValueRange`, `MajorDimension` of a compiled program; `Query`, `Scope`, `Property`, `CellRecord`,
  `QueryResponse`; `Find`, `Mode`, `SearchIn`, `Replace`). It holds no behaviour and depends only on
  `sheetrs-formula`, `serde` and `serde_json`. `sheetrs-engine` parses its requests with it (the
  messages of a wrong request are the same).
- New crate `sheetrs-query` and the executable `sheetrs query`: read the cells of an XLSX or an ODS by
  range, sheet or workbook (`--range`, `--sheet`, `--workbook`, or `--query` with the JSON of the
  protocol), with the properties asked for (`value`, `formula`, `number_format`, `display`; only
  those are read), as JSON or one cell per line (`--jsonl`). `display` is what a person sees: the
  value written with its format, in a locale (`--locale`). A sheet that does not exist is an error.
- `sheetrs query --find` keeps the cells where a text is found, as `LibreOffice`'s Find & Replace
  finds it (observed in 24.2.7): literal, regular expression (the subset of ICU that is compared
  with `LibreOffice`; anything else is refused) or wildcards (`*`, `?`, `\` to escape), match case,
  entire cells, in the value or in the formula. A search by value looks at the text of the input
  line (`1234.5` for a cell that shows `1,234.50`, `25.6%`, the date `03/15/2023` or the ISO date for
  an ISO format) and, unless `--include-formulas`, not at the result of a formula; a search by
  formula sees the notation of `LibreOffice` (`=SUM(A1,B1)`, `=$'My Sheet'.A1`).
- `sheetrs-io`: `Selection::window_everywhere` (a window for every sheet) and
  `Source::date_system`.

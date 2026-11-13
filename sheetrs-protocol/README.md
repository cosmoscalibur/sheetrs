# sheetrs-protocol

The contract of the SheetRS tools as typed JSON, with no behaviour: nothing is read, evaluated or
matched here. It depends on `sheetrs-formula` (the grammar of references), `serde` and
`serde_json`, so a client can speak the protocol without any tool.

| Module | Types |
| ------ | ----- |
| `calculate` | `Request` (`inputs`, `ranges`, `majorDimension`), `ValueRanges` (`valueRanges`: the answer of a calculation, typed, with `to_json`, `write_json` and `to_entries`), `ValueRange`, `MajorDimension` — what a compiled program (`sheetc`) answers |
| `query` | `Query` (`scope`, `properties`, `find`, `locale`, `dates`), `DateValue` (`{"type": "date"\|"time"\|"datetime", "value": "<ISO 8601>"}`), `Scope` (`"workbook"`, `{"sheet": "Data"}`, `{"range": "Data!A1:C9"}`), `Property` (`value`, `formula`, `number_format`, `display`, `comment`), `CellRecord` (the sheet is shared, `cell()` derives the A1 text from `row` and `col`), `QueryResponse` — what `sheetrs query` answers |
| `diagnostic` | `Diagnostic` (`level`, a stable `code`, `message`, `context` fields), `Level`, `LogFormat`, `Logger` (writes them on the standard error as text or JSON, with a threshold), and with the feature `cli` the flags `LogArgs` (`--log-format`, `-q`, `-v`) |
| `find` | `Find` (`text`, `mode` literal/regex/wildcards, `match_case`, `entire_cells`, `search_in` value/formula/comment, `include_formulas`, `formatted`, `filtered`) and `Replace` (a `Find` and `with`) |
| `reference` | `Reference`, a range as a caller writes it (`'My Sheet'!B2:C9`) |

Every type rejects fields it does not know, so a misspelled option is an error and not a default.

# sheetrs-protocol

The contract of the SheetRS tools as typed JSON, with no behaviour: nothing is read, evaluated or
matched here. It depends on `sheetrs-formula` (the grammar of references), `serde` and
`serde_json`, so a client can speak the protocol without any tool.

| Module | Types |
| ------ | ----- |
| `run` | `Request` (`inputs`, `ranges`, `majorDimension`), `Response` (`valueRanges` / `outputs`), `ValueRange`, `MajorDimension` — what a compiled program (`sheetc`) answers |
| `query` | `Query` (`scope`, `properties`, `find`, `locale`), `Scope` (`"workbook"`, `{"sheet": "Data"}`, `{"range": "Data!A1:C9"}`), `Property` (`value`, `formula`, `number_format`, `display`), `CellRecord`, `QueryResponse` — what `sheetquery` answers |
| `find` | `Find` (`text`, `mode` literal/regex/wildcards, `match_case`, `entire_cells`, `search_in` value/formula, `include_formulas`) and `Replace` (a `Find` and `with`) |
| `reference` | `Reference`, a range as a caller writes it (`'My Sheet'!B2:C9`) |

Every type rejects fields it does not know, so a misspelled option is an error and not a default.

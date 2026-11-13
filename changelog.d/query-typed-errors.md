### Changed

- `sheetrs query` and `query_workbook` give the `value` of an error cell as `{"error": "#DIV/0!"}`, as the answer of `sheetc calculate` does, so that an error is not taken for a text that starts with `#`.

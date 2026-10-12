### Changed

- `Sheet::hidden_rows` and `Sheet::hidden_columns` are now `Vec<(u32, u32)>` inclusive `(start, end)` ranges instead of per-index `Vec<u32>`. `extract_hidden_columns_rows_from_xlsx` and HID902's `HiddenColumnsRowsData` use the same range representation.

### Fixed

- Speed up ODS parsing of sheets ending in large repeated row blocks (e.g. LibreOffice's trailing `table:number-rows-repeated="1048575"`): remove a never-read visible-row map that cost one hash-map insert per repeated row, and store hidden rows/columns as ranges instead of one entry per index.
- Detect hidden rows written as self-closing `<row hidden="1"/>` elements in XLSX, and resolve their row index regardless of attribute order, restoring HID902 and used-range parity with ODS for hidden rows without cells.

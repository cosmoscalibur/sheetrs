### Changed

- Policy: parity of capabilities (XLSX and ODS do the same) and fidelity of behavior (the result is
  `LibreOffice`'s, on each file). A difference from `LibreOffice` exists only if the maintainer approved
  it, and it is recorded in `docs/approved-differences.md`.
- `sheetrs-io/tests/assets/oracle/make_oracle.py` records what `LibreOffice` itself says of every cell of
  the fixtures (value, shown text, format code, formula, error); `sheetrs-query/tests/libreoffice_oracle.rs`
  compares each file with its own oracle. The differences that remain are listed there as pending
  (none is approved).

### Fixed

- `ods`: a cell with content and no style takes the default style of its row or of its column
  (`table:default-cell-style-name`), as `LibreOffice` does; before, those cells had no format.
- `ods`: a negative time (`office:time-value="-PT29628H00M00S"`) is a negative number, not text.
- `sheetrs-query`: `display` of a cell whose format is `[$-409]General` was garbled; a query with
  `number_format` or `display` and no `value` returned no cell; `number_format` is `General` when the
  file says none and `BOOLEAN` for a boolean, as `LibreOffice` reports them.
- `sheetrs-edit`: a replacement always reaches the formula cells and makes them constants, as
  `LibreOffice` does; `include_formulas` is for queries (`sheetrs edit --include-formulas` is gone).

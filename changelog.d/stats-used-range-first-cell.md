### Changed

- The used area of a sheet is found once, in `sheetrs-io`, the same in an XLSX and in an ODS and for every consumer: `Sheet::used_area` is the rectangle of the used cells (a value or a formula; style-only cells, hidden rows and columns and the `<dimension>` element do not count), as the used area of LibreOffice. `Sheet::used_range` (the extent anchored at `A1`, taken from the `<dimension>` in an XLSX and from every cell in an ODS) is replaced by `used_area` and `kept_area` (every cell the file keeps, up to the last hidden row and column), both found by the same accumulator that also serves the streaming readers; the linter's REF304 and REF305 work on the latter, and their results do not change except for a worksheet without `<dimension>`, which they used to skip.

### Fixed

- `used_range` of `sheetrs stats`, `BookStats::stats` and the Python `Sheet` is the rectangle of the used cells (`D10:E11`), from the first used row and column to the last, as the used area of LibreOffice, and no longer starts at `A1`; it is the `used_area` the reader finds, with no area tallied apart. `SheetStats` has a new `used_range` field (also in the JSON output and as `Used Range` in the text output); `empty_cells` counts the cells without value or formula from `A1` to the last used cell, as before.

### Fixed

- `used_range` of `sheetrs stats`, `BookStats::stats` and the Python `Sheet` is the rectangle of the used cells (`D10:E11`), from the first used row and column to the last, as the used area of LibreOffice, and no longer starts at `A1`. `SheetStats` has a new `used_range` field (also in the JSON output and as `Used Range` in the text output); `empty_cells` counts the cells without value or formula inside that rectangle. The `<dimension>` element is not read.

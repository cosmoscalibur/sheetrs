### Added

- `sheetc` implements `CELL` (address, col, row, sheet, coord, contents, type), `INFO` (recalc, numfile and the `#N/A` types) and `STYLE`, with oracle rows computed by LibreOffice 26.8 (see `FIDELITY.md`). The types that depend on the cell's format, the file or the system stop the build.

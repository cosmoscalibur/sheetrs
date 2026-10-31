### Added

- `sheetc` implements `LINEST`, `LOGEST`, `TREND`, `GROWTH`, `FREQUENCY` and `MODE.MULT`, with oracle rows computed by LibreOffice 26.8 (see `FIDELITY.md`, *Regressions, frequencies and modes*).
- `LINEST`/`LOGEST` with statistics over a perfect fit that is not exact in the numbers, and `FREQUENCY` with equal bins, stop the build: LibreOffice answers the rounding noise of its computation, or a choice that depends on its sort.

### Fixed

- `INDEX` of an error is that error (it was `#REF!`).

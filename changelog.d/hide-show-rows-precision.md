### Fixed

- The default mode of `sheetrs edit` (the cache) calculates, when rows are hidden or shown, only for the rows whose visibility changes and that hold a value or a formula, as `LibreOffice` does: hiding a row that is hidden, showing one that is visible, or hiding or showing a row with no cell no longer marks the formulas that read it as stale. The input is read for the hidden rows of the sheets named, as a stream that stops after the last row asked for (`Book::row_state`).

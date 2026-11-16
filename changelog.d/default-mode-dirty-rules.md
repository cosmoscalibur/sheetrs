### Fixed

- The default mode of `sheetrs edit` (the cache) calculates what `LibreOffice` calculates when it makes the same deletion and the file is
  saved and opened again, and keeps the result the file caches for the rest, in both formats: a formula that reads through a name is
  calculated only when the deletion moved, shrank or removed the target of the name (or a name was deleted); every formula is calculated
  when a sheet is deleted; `ROW` and `COLUMN` only when their cell, or what they are given, moved; a formula that already held an error
  `#REF!` keeps its result. A formula is no longer judged by the cell at its index when the compiler takes some formulas from the cache
  (the results of the formulas after one that was taken were misread).
- The formulas that are not reached keep their result however many are (an earlier version calculated everything where more than a quarter
  were stale or had no result).

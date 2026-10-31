### Added

- `sheetstats` reports the end of each sheet (`GV5439`, the corner of its used area) and its comments, charts and tables, with the workbook totals, in XLSX and ODS. The readers give each sheet its `comments`, `charts` and `tables`.

### Changed

- `sheetstats` counts the cells while it reads them instead of keeping the workbook: the peak memory on the 39 MB production XLSX falls from 462 MB to 22 MB and on the ODS from 440 MB to 5 MB, with the same counts.

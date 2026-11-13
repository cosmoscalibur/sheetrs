### Added

- Python: `Book.to_bytes()` writes the workbook the accumulated operations make to memory, as `bytes` in the format of the source (the same bytes `save(path)` writes for the same plan); with no operations it returns the workbook as it is.

### Changed

- Tests: no test judges an XLSX by an ODS or the other way round. The oracle of an XLSX is what LibreOffice gives when it opens and recalculates that file (`tests/assets/semantics*.xlsx.recalc.json`, made by `scripts/gen_semantics_fixture.py`: the code of each error, which an XLSX caches as `#VALUE!`, and the rows its export rewrites); the cross-format plan test and `scripts/compare_formats.py` are gone, and three tests are renamed for what they check on each file.

### Fixed

- `OFFSET` with an error as its reference gives that error (`OFFSET(NA(),0,0)` is `#N/A`, `OFFSET(1/0,0,0)` is `#DIV/0!`, an XLSX `OFFSET(#REF!,0,0)` is `#NAME?`), as LibreOffice does; the engine gave `#VALUE!`.

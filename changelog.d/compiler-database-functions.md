### Added

- `sheetc` implements the database functions of LibreOffice, `DSUM`, `DCOUNT`, `DCOUNTA`, `DAVERAGE`, `DGET`, `DMAX`, `DMIN`, `DPRODUCT`, `DSTDEV`, `DSTDEVP`, `DVAR` and `DVARP`, with 970 oracle rows computed by LibreOffice 26.8 (see `FIDELITY.md`, *Database functions*).
- `DCOUNT` and `DCOUNTA` with field 0 stop the build: LibreOffice's count is not a function of the records that match.

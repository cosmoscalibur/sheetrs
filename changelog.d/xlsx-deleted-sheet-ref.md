### Fixed

- Deleting a sheet of an XLSX leaves the bare `#REF!` that LibreOffice and Excel write, and the results are what LibreOffice shows when it opens the file it saves after the same edit (`COUNT` 0, `COUNTA` 1, `IFERROR` catches the reference; `SUM` and the rest are `#NAME?` in an XLSX, after a deleted sheet as after deleted cells): `sheetrs-edit/README.md` says what can and cannot be reconciled.

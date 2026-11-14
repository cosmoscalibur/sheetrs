### Fixed

- Deleting a sheet of an XLSX leaves `#REF!` as the cached result of the formulas that read it, as the edit does in LibreOffice 26.8 (it was `#NAME?` in the formulas of a calculation, and 0 in `COUNT(Data!A:A)`); the formula text is the bare `#REF!` that LibreOffice and Excel write. LibreOffice reads that text back as an unknown name, also in the file it saves itself: `sheetrs-edit/README.md` says what can and cannot be reconciled.

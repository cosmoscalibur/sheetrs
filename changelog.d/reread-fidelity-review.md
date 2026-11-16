### Fixed

- The functions that read a deleted sheet in an XLSX (`SUM`, `AVERAGE`, `MAX`, `COUNTIF`, `INDEX`... and a calculation over the reference) give `#NAME?`, as `LibreOffice` shows when it opens the file it saves after the same edit (the cells that deleted cells left already did): the approved difference that gave `#REF!` followed the memory of `LibreOffice`, which is not the reference, and is gone with the code that existed for it (`Removed`, `RemovedAt`, `RemovedRefs`). The definition of a defined name that went with the sheet reads as `#NAME?` too.
- A reference that is `#REF!` in any spelling (`COUNT(Data!#REF!)` of an XLSX, as Excel writes it) is read by the compiler as the plain error `#REF!` as a whole, so its values are those of the file `LibreOffice` saves and opens again (`COUNT` 0, `COUNTA` 1, `IFERROR` catches it) with no text rewritten in the passes that move nothing (hide, replace, calculate); the formula text of an untouched cell keeps `Data!#REF!`.

### Changed

- `docs/approved-differences.md` lost the entries that the rule of the re-read made unnecessary (the spelling of a reference that went, the deleted-sheet `#REF!`, the whole column over several sheets, the cache that was not kept) and lists, as proposed, the cases where the default mode shows other values than `LibreOffice` after an edit with no `calculateAll()`.

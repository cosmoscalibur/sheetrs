### Fixed

- An ODS formula with `[#REF!]` (what LibreOffice writes for a deleted reference) reads as `#REF!`; it was read as a reference to the external workbook `[1]`, which also listed a workbook that does not exist.
- The formula parser reads `Sheet!#REF!`, the form that a deleted cell leaves in an XLSX.
- An ODS formula with a reference across sheets (`[$T.A2:$V.A4]`) reads as `T:V!A2:A4`, the form an XLSX writes, where it read as `T!A2:V!A4`. The parser reads `T:V!A:A` as the 3D reference from sheet `T` to sheet `V` even when the names look like columns.

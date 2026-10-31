### Fixed

- DATA702 (number stored as text) no longer reports the text a formula returns (`=IF(A1,"0","1")`, `CONCAT` of digits): only constants are stored data, as in DATA706. Reading the cached string result of a formula as text (it was read as a number) had made 108 such ranges appear in the large production workbook.

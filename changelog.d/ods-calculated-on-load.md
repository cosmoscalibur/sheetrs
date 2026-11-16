### Changed

- The default mode (the cache) calculates what `LibreOffice` calculates when it opens the file, in both formats, and keeps the result the file caches for
  the rest. In an ODS those are the formulas whose cached result is a number, a currency or a date in a cell with no number format (not a
  percentage, a boolean or a time), a cached error, an array formula and a result with no type; what reads them is not calculated, as in `LibreOffice`;
  the writer of an ODS no longer needs the numbers to be wrong to be calculated: the conversion of `book.ods` writes 14 and 8, as `LibreOffice` does (it wrote
  the cached 140 and 99), and the difference that was approved for it is gone. In both formats `HYPERLINK`, every type of `CELL`, `INFO` and `FORMULATEXT` are
  calculated on load, with what reads them, and so are `INDIRECT` and `OFFSET`, which no longer make the workbook calculated whole (the formulas that do not
  depend on them keep their result).
- Hiding or showing rows calculates a `SUBTOTAL` that reads a range with one of them and the formulas of a group (adjacent formulas of a column with
  the same relative references) that read it, as `LibreOffice` does; `AGGREGATE` and a formula alone keep their result.
- `ISFORMULA` and `FORMULATEXT` of a cell that a deletion removed give what `LibreOffice` gives (`FALSE`; `#NAME?` in an XLSX and `#REF!` in an ODS) instead of an
  error.
- `write_to_libreoffice.rs` runs the default mode now against `LibreOffice` for every fixture in both formats (the ODS ones were left out), with the
  new fixtures `classes` (a formula of each class of function) and `indirect`; the probes are in `sheetrs-edit/tests/assets/write_to`.
- The defined names of an ODS are read from the skeleton of `content.xml` (its tables and the blocks of names, searched in the bytes) when they are asked alone,
  which an edit that deletes rows or columns does before its pass.

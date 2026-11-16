### Added

- `sheetrs-io/README.md`, *What the writers change of a result when they write it*: a table of what each kind of result (errors, numbers of every size, texts, booleans, dates, times, percentages, currencies, NaN, infinities, empty cells, arrays, in both date systems and both formats) becomes in the file when a writer writes it, against what `LibreOffice` writes and shows after saving and opening again, each row proven by the oracle `sheetrs-edit/tests/assets/writer_results` (`make_oracle.py`, `oracle.json`) and by `sheetrs-edit/tests/writer_results.rs`. A conversion of formulas to values leaves an empty text where the formula gave an error, as `LibreOffice` does (`.uno:ConvertFormulaToValue`, re-read, every kind of error, both formats): that is the reference, not a loss. The differences found (numbers keep all their digits where `LibreOffice` keeps 15, infinities of an XLSX, the empty text constant of an ODS, the text format of constants, the form of the cells) are proposed, not approved, in `docs/approved-differences.md`.

### Fixed

- In an ODS, a formula that is a single decimal number (`=1.5`, `=0.25`, `=12345.678`) was read as a reference to a cell (`1.5` as the cell `5` of a sheet `1`), so a conversion of formulas to values that calculates them (`calculate_all`, and the cells an ODS calculates when it is opened) wrote `5`, `25` and `678` for them.

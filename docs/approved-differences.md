# Approved differences from LibreOffice

SheetRS gives the result LibreOffice gives. A difference exists only if the maintainer approved it,
and it is recorded here with where it is described. Anything else that differs is a bug or a pending
decision, never a documented feature.

| Difference | Where | Why |
| ---------- | ----- | --- |
| A query has no direction (rows or columns): the cells come in the order of the sheet | `sheetrs-query/README.md` | Scope of the tool |
| Some mathematical functions are computed with `libm` and may differ in the last digits from LibreOffice's | `sheetrs-compiler/docs/FIDELITY.md` | Technical |
| `LONGTEXT` (the pieces of a text over 255 characters) is a feature that is enabled on request | `sheetrs-compiler/README.md` | Scope |
| `include_formulas` is `false` by default in a query: a search by value looks at what was typed, not at the results of formulas | `sheetrs-query/README.md` | Usability |
| A cached error of an XLSX is read as the file keeps it (`#VALUE!`), where LibreOffice recalculates the cells with an error when it opens the file and shows its own code (`Err:502`, `Err:522`). To have LibreOffice's errors, recalculate the workbook first (see the recipe) and read the result | `sheetrs-query/README.md`, `docs/recipes.md` | Reading does not calculate |
| Regular expressions outside the subset of ICU that was compared with LibreOffice are refused in a search and in a replacement, instead of giving another result (issue #82 extends the subset) | `sheetrs-query/README.md` | Scope |
| Only the three date systems of the LibreOffice Calc interface are supported: 1899-12-30 (the default, and the 1900 system of an XLSX), 1900-01-01 (StarCalc 1.0) and 1904-01-01. LibreOffice can store another null date (its API, a crafted ODS `table:null-date`); SheetRS gives an error that names the date, when the workbook is read, in every tool, instead of numbering the dates from it | `sheetrs-compiler/docs/FIDELITY.md` (Date system), `sheetrs-io` (`Error::UnsupportedDateSystem`) | Scope: an arbitrary null date is not reachable from the interface and no tool writes it |
| `FORECAST.ETS` and its variants, with an error written in an option or next to a series of a single value: LibreOffice reads the error as a number (the payload of the NaN) and uses it without checking it (`FORECAST.ETS(7,{1;3;2;5;4;6},{1;2;3;4;5;6},#NAME?)` is 9.22337203685478E+18, the largest 64-bit integer), SheetRS answers the error | `sheetrs-compiler/docs/FIDELITY.md` | An unchecked conversion of LibreOffice, not a rule; the error is the useful answer |
| `SUM(Data!A1:Wide!A1048576)` in an XLSX (a whole column over two sheets): the headless LibreOffice that recalculates the file gives `#NAME?`; SheetRS reads the form and adds the cells, as Excel does and as LibreOffice itself writes the form when it saves | `sheetrs-compiler/docs/FIDELITY.md` | The form is written by LibreOffice and by Excel; the `#NAME?` is a limit of the headless mode |
| Only the A1 notation of references is supported. LibreOffice can read and show references in R1C1 notation (*Tools ▸ Options ▸ Calc ▸ Formula*, and in the API); SheetRS has no R1C1 grammar: a range written in R1C1 (`T!R1C1:R2C2`) is refused in every tool with "is not a range", and the formulas of a file are always read in the notation the format stores (A1) | `docs/coding-patterns.md` (Formula Normalization) | Scope: R1C1 is not usual and no file stores it |

Adding a line requires the maintainer's approval, in the pull request or the issue it came from.

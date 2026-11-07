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

Adding a line requires the maintainer's approval, in the pull request or the issue it came from.

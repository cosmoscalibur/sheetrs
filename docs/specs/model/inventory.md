# Macros and external links

The model lists the macro code and the documents a workbook refers to; it never runs the code
and never opens the documents.

## Behaviour of the model

- `Workbook::macro_modules` lists each module as the library (or project) it is in, its name and
  its kind, `Basic` or `Vba`. The code of a module is not read.
- `Workbook::external_links` lists each document once, with the target as the file writes it. A
  path relative to the workbook stays relative, because a reader of a stream does not know where
  the workbook is.

## Observations

| Behaviour | Test |
| --- | --- |
| LibreOffice names the documents a workbook refers to and the modules it holds, which the oracle reports for every case after the workbook is loaded. | every `..._over_xlsx`, `..._over_ods` and `..._over_csv` test of the readers |

## Open defects

- The security warning that the plan asks for, and the exposure of the inventory to the linter,
  do not exist yet.
- That the calculation behaves as if a macro did not exist has not been compared with the oracle.

### Fixed

- Fix `sheetrs query --find --include-formulas` (and `Book.find` with `include_formulas`) not finishing on a workbook with millions of formulas: the inference of the format of each result read the file again for every different range a formula refers to, and looked for the cells of each range through all the cells read. It reads each sheet whole after eight windows and finds the cells of a range by position; the 2.4-million-formula workbook is searched in 18 s (XLSX) and 23 s (ODS), where it had not finished in five minutes. The answers do not change.

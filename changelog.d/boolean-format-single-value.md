### Fixed

- A formula that ends with a function that keeps the boolean format of its argument (`VALUE(TRUE())`, `NUMBERVALUE`, `+TRUE()`) gives the boolean LibreOffice shows, not the number 1; a function or an operator that sets a number format (`ABS(TRUE())`, `TRUE()+0`, `-VALUE(TRUE())`) still gives the number. `ISLOGICAL(-TRUE())` is true. The table of the functions comes from `sheetrs-compiler/scripts/probe_logical.py`.

### Fixed

- A plan that hides rows and deletes earlier rows in the same write (`hide_rows("Data!5")` with `delete_rows("Data!1")`) calculates the `SUBTOTAL` that reads the hidden row where the row now is, as `LibreOffice` does when the saved file is opened again: the rows hidden were kept with the numbers of the input while the formulas had the ones of the output, so an XLSX kept the cached result of a `SUBTOTAL` that the hidden row changes.

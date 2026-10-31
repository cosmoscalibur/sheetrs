### Fixed

- XLSX and ODS: a numeric cell (`t="n"`, `office:value-type="float"`) in a cell formatted as Text is read as a number, as Excel and LibreOffice keep it, instead of as text. Lookups and sums over such cells (a list of codes formatted as Text) now see numbers.

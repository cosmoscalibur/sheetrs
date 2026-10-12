### Fixed

- Fix XLSX formulas, shared strings and defined names dropping escaped characters (`"`, `<`, `>`, `&`, `'` and character references) and the spaces next to them. Formulas such as `IF(A1<0,"Y","N")` were read as `IF(A10,Y,N)`, causing wrong results in formula rules and differences with the ODS reader. ODS cell text and date formats resolve references the same way.

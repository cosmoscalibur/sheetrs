### Changed

- The cells of an XLSX sheet and the rows and cells of the `content.xml` of an ODS are read with a lexer of the suite (`sheetrs_io::xml::Lexer`) instead of the events of `quick-xml`: reading a workbook is faster (`sheetrs stats` of the 2.4 M formulas of the benchmarks, XLSX and ODS, and every tool that reads a workbook with it) at the same memory, and what is read is the same (the tests compare the lexer with `quick-xml` on every XML part of the test workbooks, and the readers with the previous ones on every workbook of the repository). The properties, styles and names parts, the rare elements of a sheet and the writers keep the events of `quick-xml`.
- `parse_cell_ref` no longer allocates a string for each reference it reads.

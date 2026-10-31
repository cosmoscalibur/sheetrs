### Fixed

- ODS: the comment (`office:annotation`) of a cell is no longer read as the cell's value. An empty cell with a comment used to become a text cell holding the comment, which changed the type seen by rules and by anything that reads the cell.
- ODS: runs of spaces (`<text:s/>`), tabs and line breaks inside a cell's text are kept, and the paragraphs of a cell are joined with a line break, so cell texts equal what the spreadsheet application shows.
- ODS: text inside a formula's string literal is no longer rewritten as a reference (`"[$$-409]#,##0"` lost its brackets and `$`), which changed the format codes of `TEXT` and any text that looked like a reference.

### Added

- `find` (`sheetrs query --find --regex`, `Book.find` in Python and WebAssembly) and `replace` (`sheetrs edit --replace --regex`, `Book.replace`) read the anchors `^` and `$` of a regular expression as LibreOffice does in a cell: the start and the end of its text (`$` also before the line terminator that ends it), not of its lines (issue #82). Both read a pattern that can match the empty text (`x*`, `^`, `$`, `a*$`, `(^|$)`) as LibreOffice does: `find` reports the cells that `findAll` reports (the ones `replace` would change), and the replacement is made in the text itself and the search goes on after it, an empty match is replaced only at the very end of the text (`abc` with `$` replaced by `X` is `abcX`; `^` and `x*` change nothing), and no search begins at the end of the text. Compared with `findAll` and `replaceAll` of LibreOffice 26.8 in XLSX and ODS. A search that is still outside the subset of ICU (`\b`, look-around, a repetition of something that can match the empty text) is refused with the list in `sheetrs-query/README.md`.
- The empty search text matches nothing in `find` and in `replace`, in every mode (literal, regex, wildcards), as in LibreOffice: no cell is found or replaced, with no error and before the workbook is read or the pattern compiled (it was the refusal `unsupported search`), in the command line, Rust, Python and WebAssembly.

### Fixed

- An ODS cell whose text has an empty paragraph (`a` and an empty line, or an empty line and `a`) is read with that line break, as LibreOffice reads it.

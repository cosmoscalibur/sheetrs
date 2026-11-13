### Added

- `find` (`sheetrs query --find --regex`, `Book.find` in Python and WebAssembly) reads the anchors `^` and `$` of a regular expression as LibreOffice does in a cell: the start and the end of its text (`$` also before the line terminator that ends it), not of its lines (issue #82). `replace` keeps refusing them.

### Fixed

- An ODS cell whose text has an empty paragraph (`a` and an empty line, or an empty line and `a`) is read with that line break, as LibreOffice reads it.

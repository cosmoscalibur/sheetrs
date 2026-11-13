### Added

- Notes: `sheetrs query -p comment` and `--in comment` read and search the notes of XLSX (legacy comments, rich runs joined) and ODS (`office:annotation`), and `sheetrs edit --replace --in comment` replaces in them as LibreOffice's `replaceAll` does. `Properties::comment`, `CellSink::comment`, `Property::Comment`, `SearchIn::Comment` and `CellRecord::comment` are the library surface. Checked against notes LibreOffice made in both formats.
- A wildcard search whose text ends in `*` takes the rest of the text, as LibreOffice's does.

### Changed

- Every cell of an array formula, the interior ones too, reports its formula (`{=…}` in a search) in both formats, as LibreOffice shows them. The statistics count them as formulas.
- Writing a recalculated workbook (`sheetc calculate --recalculate --output`, `fill_template`, `Editor::write_cells`) deflates the parts it rewrites at level 4 instead of the default: 4 million cells, 36.5 MB output, the write takes 8.2 s instead of 11.4 s (the whole run, 22.6 s before, about 20 s).

### Changed

- `sheetrs edit` on an ODS starts with the calculation that `LibreOffice` runs when it opens one, in the mode of the run: the formulas it calculates on load by default (a cached number in a cell with no format, an error, no result, a few functions), every formula with `--calculate-all`; the roots are the only difference between the modes at the start. The startup is the first pass that runs the engine (a pass of its own when nothing else in the plan needs it, as hiding a column), and the roots are found in the read that compiles the workbook: the scan of `content.xml` that decided before whether to calculate (9.3 s on 2.4 M formulas) is gone. After the startup each operation applies its own rule. An edit that leaves nothing stale now shows what `LibreOffice` calculates on load (`rate*2` 14, `ROW()` 8 where the file caches 140 and 99) and keeps a cached text, and the explanation (`--dry-run`) says the startup first. An XLSX has no startup.
- A conversion of formulas to values is one pass in every mode where the engine runs for it: the engine calculates, in the pass that writes the conversion, what is stale (the cache) or everything (an ODS with `--calculate-all`), and the rest of the scopes is converted with the result the file caches; no pass that writes every result comes first.

### Fixed

- Hiding or showing a row calculates, in an ODS, a `SUBTOTAL` that reads a single cell of that row, as `LibreOffice` does (an XLSX keeps its result); the `visibility` fixture compares both formats.

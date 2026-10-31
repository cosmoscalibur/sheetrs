### Changed

- REF310 shows the whole block a range lacks and the replacement range written like the original (`$` flags and sheet qualifier kept): "Range INGRESOS!$H$61:$Q$61 may be too short: the block continues in INGRESOS!$R$61:$AA$61 (10 columns). Suggested range: INGRESOS!$H$61:$AA$61."

### Added

- `sheetc`: the `enable_features` configuration key (`--enable-features`), with the feature `LONGTEXT` (off by default): `_xlfn._LONGTEXT(…)` with constant text arguments (how LibreOffice 26.2 wrote a literal over 255 characters) is joined into one text literal when the program is built.
- ERR102 does not count a saved `#NAME?` as an error when the formula calls a function with the `_xlfn.` prefix (the application that saved the file did not know it), nor a `#NAME?` that a cell inherits from one. `_xlfn._LONGTEXT(…)` is not an Excel function, so it stays an error unless the new global `enable_features` list names `LONGTEXT`.

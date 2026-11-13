### Added

- `sheetrs-query`: a formula cell whose format is `General` takes the number format of its result as LibreOffice does (`number_format` and `display`): `=A1` over a date, `=A1+1`, `=A4*2` over a currency, aggregates, `IF`, `INDEX`, `INDIRECT`, `OFFSET`. The rules were obtained by probing LibreOffice 26.8 and are checked against it with the fixture `inherit.{xlsx,ods}`. In an XLSX only the volatile cells and those that depend on them are interpreted at load, so only those change.

### Fixed

- `General` in a cell is written in the scientific notation below `1E-4` as LibreOffice does (`4.44444444444444E-05`, `1E-10`) with a two-digit negative exponent; the text a number becomes inside a formula is not changed.
- The oracle test no longer tells apart `[$EUR-407]` from `[$EUR]`: the language mark of a currency is notation.
- `scripts/libreoffice-server.sh` no longer reports a server that is not running (the check matched itself).

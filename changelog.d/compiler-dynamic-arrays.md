### Added

- `sheetc` implements more functions of LibreOffice with oracle rows computed by LibreOffice 26.8: `SEQUENCE`, `FILTER`, `SORT`, `SORTBY`, `TAKE`, `DROP`, `CHOOSECOLS`, `CHOOSEROWS`, `HSTACK`, `VSTACK`, `EXPAND`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`, `RANDARRAY`, `XMATCH`, `TEXTSPLIT`, `TEXTBEFORE`, `TEXTAFTER`, `ENCODEURL` and `ROT13` (see `FIDELITY.md`, *Dynamic arrays*).
- `SORTBY`, and `SORT` of an array, by keys that have an error stop the build: the order LibreOffice gives is not a function of the keys.

### Fixed

- `MATCH`, `VLOOKUP`, `HLOOKUP` and `XLOOKUP` with a text criterion over a range now find the number the text shows (`MATCH("5", A1:A8, 0)`), wildcards included, as LibreOffice does.
- `ROWS` and `COLUMNS` of an error give the error.
- An XLSX formula with both prefixes (`_xlfn._xlws.FILTER`) is no longer `#NAME?`.

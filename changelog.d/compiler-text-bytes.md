### Added

- `sheetc` implements `LENB`, `LEFTB`, `RIGHTB`, `MIDB`, `FINDB`, `SEARCHB`, `REPLACEB`, `ASC`, `JIS`, `BAHTTEXT` and `REGEX` with oracle rows computed by LibreOffice 26.8 (see `FIDELITY.md`, *Byte functions*). The width table of the byte functions and the conversions of `ASC` and `JIS` are generated from LibreOffice (`scripts/gen_width_data.py`).
- `REGEX` takes the same subset of ICU's syntax as the search of the criteria functions, now with its capturing groups (`$1`…); an expression outside it, and a byte function that cuts through half of a character above the plane, stop the build.

### Changed

- Replace `bincode`, which is unmaintained (its 3.0.0 release is a tombstone that does not compile), with `postcard` as the encoding of the program and of the blocks of constants inside an executable built by `sheetc build`: the executable of a workbook without declared inputs and outputs is about half the size (`data.xlsx` 207 → 115 MB, `formulas.xlsx` 194 → 94 MB) and starts with less memory.
- Update `zip` to 9 (a part name that is not valid text is skipped when reading and is an error when writing), `sha2` to 0.11, `pprof` to 0.15 and `serde_json` to 1.0.154.

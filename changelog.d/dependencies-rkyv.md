### Changed

- Replace `bincode`, which is unmaintained (its 3.0.0 release is a tombstone that does not compile), with `rkyv` as the encoding of the program and of the blocks of constants inside an executable built by `sheetc build`. The archive is validated before it is read.
- Update `zip` to 9 (a part name that is not valid text is skipped when reading and is an error when writing), `sha2` to 0.11, `pprof` to 0.15 and `serde_json` to 1.0.154.

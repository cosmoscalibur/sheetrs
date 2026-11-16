### Added

- The payload of a compiled program (an executable and a WebAssembly module) carries a schema number next to its magic (`bundle::SCHEMA`),
  written by `sheetc` and checked by the runtime before it decodes the program: a payload of another number is refused with a message that
  names both (`payload schema 0 ... reads schema 1`) instead of being decoded as the layout the runtime knows (or reported as "no program
  embedded"). A change of the order or of the type of a field of `Program` or of what it holds bumps it (a test pins the layout).
- `sheetc build` refuses a runtime stub (`--runtime`) and a base module of another payload schema or of another ABI version, naming both; a
  base or a stub that says nothing is refused too. The base module and the stub carry what they read (`SHEETRS_INFO`, the schema marker).
- The module has the export `sheetrs_payload_schema`, added without bumping `sheetrs_abi` (the ABI version is an integer bumped only on an
  incompatible change of the exports or of what they mean; an added export is found by looking for it). The host README says that a module is
  only valid with the `sheetc` that built it.

### Changed

- `bundle::read_embedded` returns `Result<Option<Embedded>, PayloadError>` and `bundle::read_payload` returns `Result<Embedded, PayloadError>`.

### Changed

- `sheetlint.toml` lists every configuration key with its default value from the code, and its comments use current rule IDs.

### Fixed

- Fix the `sheetlint` rule reference: statuses, rule IDs and descriptions now match the registry, the Params column lists every parameter each rule reads, a status legend is added, and unregistered rule IDs are removed (#50).
- Add a test that fails when the rule reference or the `sheetlint.toml` example drifts from the registry.

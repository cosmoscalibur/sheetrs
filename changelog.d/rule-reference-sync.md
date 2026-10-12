### Changed

- The `sheetlint` rule reference drops the Status column: every listed rule is implemented and active by default.
- `sheetlint/comparative_report.md` drops the Support Status and Migrated columns, corrects scopes, examples and rule dependencies against the code, and adds a Parity column with the differences from PerfectXL Risk Finder and ExceLint.
- `sheetlint.toml` lists every configuration key with its default value from the code, and its comments use current rule IDs.

### Fixed

- Fix the `sheetlint` rule reference: rule IDs and descriptions now match the registry, the Params column lists every parameter each rule reads, and unregistered rule IDs are removed (#50).
- Add a test that fails when the rule reference or the `sheetlint.toml` example drifts from the registry.

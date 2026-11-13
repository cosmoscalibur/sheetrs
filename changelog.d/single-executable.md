### Changed

- One executable, `sheetrs`, with the subcommands `query`, `edit`, `lint` and `stats`, replaces `sheetquery`, `sheetedit`, `sheetlint` and `sheetstats`. `sheetc` keeps `build`, `calculate`, `inspect` and `verify`. The release archives carry `sheetrs`, `sheetc` and `sheetrs-rt`. Each tool crate holds its subcommand behind the `cli` feature (`sheetrs_query::cli`, `sheetrs_edit::cli`, `sheetrs_linter::cli`, `sheetrs_stats::cli`: the clap `Args` and `run`); with the feature off the libraries do not depend on clap, and neither do `sheetrs-wasm` nor `sheetrs-py`. The SARIF `tool.driver.name` is still `sheetlint`.
- `sheetc calculate` and the compiled executables take the request as `--request JSON|@FILE|-` and ranges as `--range RANGE` (repeatable); `--cinp` and `--cout` are gone there. `cinp` and `cout` name only the declaration of inputs and outputs when a program is built.

### Added

- `sheetrs query --list-sheets` (the names of the sheets, in order) and `--list-names` (the defined names and what they refer to), as JSON.

### Removed

- `sheetrs edit --list-names`: `sheetrs query --list-names` replaces it.

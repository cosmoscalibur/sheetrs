### Fixed

- Keep `zip` below 9.1. From 9.1 it depends on `getrandom` with `wasm_js` on wasm32 whatever its features: the WebAssembly modules of `sheetc` (`sheetrs-rt-template.wasm`) asked the host for five imports of the `wasm-bindgen` glue where they have none, so `sheetc build --target wasm --embed-template` produced modules that no plain host could instantiate, and `wasm-bindgen` was built without its `std` feature, which fails in a release build on Rust 1.99 (the job `wasm npm package` of the CI was red since the update).

### Changed

- The tests of `sheetc` that mixed several behaviors or had vague names are split and have short names that say the behavior (`locale_flag_is_refused_at_run_time`, `info_lists_inputs_and_outputs`, `explicit_config_replaces_the_discovered_one`, ...). Two checks that only fixed a design decision (that there is no `--dialect` flag, and that `outputs` is no longer a key of `sheetc.toml`) are removed: an unknown flag and an unknown key are already refused by their own tests.
- `CONTRIBUTING.md` (*Test Patterns*): a test is named for the behavior it checks, tests behavior and not design, and checks one behavior.

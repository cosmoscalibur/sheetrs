### Fixed

- Keep `zip` below 9.1. From 9.1 it depends on `getrandom` with `wasm_js` on wasm32 whatever its features: the WebAssembly modules of `sheetc` (`sheetrs-rt-template.wasm`) asked the host for five imports of the `wasm-bindgen` glue where they have none, so `sheetc build --target wasm --embed-template` produced modules that no plain host could instantiate, and `wasm-bindgen` was built without its `std` feature, which fails in a release build on Rust 1.99 (the job `wasm npm package` of the CI was red since the update).

### Changed

- The tests of `sheetc` that had vague names have short ones that say what they check (`locale_flag_is_refused_at_run_time`, `emitted_ir_calculates_without_the_workbook`, ...). The two that checked several things are split (`run_time_options_are_only_what_a_file_cannot_say` in three, `executable_is_self_describing_and_strict` in two), and the test that only said that there is no `--dialect` flag is removed: an unknown flag is already refused by `bad_requests_are_refused`.

### Fixed

- Keep `zip` below 9.1. From 9.1 it depends on `getrandom` with `wasm_js` on wasm32 whatever its features: the WebAssembly modules of `sheetc` (`sheetrs-rt-template.wasm`) asked the host for five imports of the `wasm-bindgen` glue where they have none, so `sheetc build --target wasm --embed-template` produced modules that no plain host could instantiate, and `wasm-bindgen` was built without its `std` feature, which fails in a release build on Rust 1.99 (the job `wasm npm package` of the CI was red since the update).

### Changed

- The tests of `sheetc` that had vague names say what they check (`the_locale_is_fixed_when_the_program_is_built_and_a_run_time_flag_is_refused`, `the_emitted_ir_calculates_without_the_workbook`, ...), and the two that checked several things are split: `run_time_options_are_only_what_a_file_cannot_say` in four and `executable_is_self_describing_and_strict` in two.

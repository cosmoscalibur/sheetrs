### Added

- Add a CI workflow that runs on every pull request and on pushes to `main` and `develop`: `cargo fmt --check`, `cargo clippy -D warnings`, `cargo test`, the `sheetrs-wasm` WASM build and the `sheetrs` build without default features. Dependabot also keeps the workflow's actions up to date.

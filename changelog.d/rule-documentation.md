### Changed

- Every linter rule module documents what it reports, what it does not (and why), its parameters and its severity, in the layout described in the `sheetrs_core::rules` module; a test checks it. DATA702 documents that it reports stored constants only. CI builds the documentation with `RUSTDOCFLAGS="-D warnings"` (`cargo doc --workspace --no-deps --document-private-items`).

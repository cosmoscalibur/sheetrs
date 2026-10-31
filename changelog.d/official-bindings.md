### Added

- Add `sheetrs-py`, the official Python bindings (PyO3, abi3 wheels for Python ≥ 3.10): `lint`,
  `stats`, `edit` and `rules`, from a path or from bytes, with type stubs and a pytest suite.
- Add `sheetrs_core::report` (`ViolationRecord`, `violation_records`): the flat, serializable form
  of the violations, shared by the bindings.
- Add `lint_workbook_records` and `workbook_stats` to `sheetrs-wasm`: structured results besides
  the text reports. Add `scripts/package-wasm.sh`, which builds the npm package.
- Add a release workflow that attaches executables (Linux x86_64/aarch64, macOS arm64, Windows
  x86_64), Python wheels and sdist, and the npm tarball to each `vX.Y.Z` release; CI builds the
  Python extension and the npm package on every pull request.

### Changed

- `sheetrs-stats` is a library as well as the `sheetstats` executable (`stats::WorkbookStats`,
  `StatsSink`).
- Path dependencies between the publishable crates carry a version, and the crates have keywords and
  categories, so they can be published to crates.io. `sheetrs-wasm` is not published.

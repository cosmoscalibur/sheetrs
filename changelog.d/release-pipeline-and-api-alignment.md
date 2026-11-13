### Added

- `docs/api-alignment.md`: the rules that keep the command line, Rust, Python and WebAssembly aligned and the table of what each operation is called in every layer.
- Release workflow: a manual run is a dry run that builds and smoke-tests every artifact without creating a release; a smoke job runs the executables, installs the wheel in a clean environment and opens a workbook with the npm tarball; the release notes start with the section of the version in `CHANGELOG.md`.
- CI: a job that builds `sheetrs` and `sheetc` with `--locked` and runs them on a real workbook.
- README: installation from GitHub with `uv`, `pip` and `cargo`, and the npm tarball of a release.

### Changed

- The release builds the Linux executables only (macOS and Windows are issue #99); the package no longer claims macOS and Windows as operating systems.

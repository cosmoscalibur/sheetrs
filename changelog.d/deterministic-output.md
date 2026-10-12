### Added

- Add `--baseline <ref>` to `scripts/benchmark.sh`, which also times a git ref in the same session. By default the script times only the current tree.

### Changed

- REF310 compares value types cell by cell between the range's last row (or column) and the first adjacent line, an empty cell being its own type, instead of using the range's dominant type. Every distinct range in a formula is evaluated and reported once.
- Violations sort by scope, then rule ID. CPX504–509 and CPX507 emit their violations at sheet end, in cell order. Production benchmark, with these REF310 changes: about 0.5% fewer user-space instructions, peak RAM up to 0.6% higher.

### Fixed

- Fix `sheetlint` output order varying between runs of the same build (#33).
- Fix REF310 reporting a different set of violations between runs when value types tied in a range (#34).
- Fix REF310 skipping later ranges of a formula once its first range was suppressed.

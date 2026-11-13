## Description

<!-- Describe what this PR does and why. -->

Closes #<!-- issue number -->

## Changes

<!-- List the key changes in this PR. -->

-

## Testing

- [ ] `cargo test --workspace --exclude sheetrs-py` passes (and `pytest sheetrs-py` if the Python binding changed)
- [ ] `cargo clippy --workspace --exclude sheetrs-py --all-targets -- -D warnings` passes
- [ ] `cargo fmt --check` passes
- [ ] Checked against LibreOffice (its result on each file), in XLSX and ODS
- [ ] No difference from LibreOffice without the maintainer's approval (see `docs/approved-differences.md`)

## Changelog

- [ ] Added a changelog fragment in `changelog.d/`
  (see [CONTRIBUTING.md](CONTRIBUTING.md#changelog-fragments))

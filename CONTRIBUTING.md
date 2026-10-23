# Contributing to SheetRS

Thank you for your interest in contributing to the SheetRS Suite.

## Development Setup

**Prerequisites**: Rust ≥ 1.99 (edition 2024).

```bash
git clone https://github.com/cosmoscalibur/sheetrs.git
cd sheetrs
cargo build
```

## Testing

### Running Tests

```bash
# Run all tests (debug mode — safe when not all rules are enabled)
cargo test

# Run tests in release mode (recommended when all rules are enabled)
cargo test --release
```

### ODS/XLSX Parity

Feature and behavior parity between ODS and XLSX formats is a requirement. Any
difference detected should be considered a potential bug.

Changes should be tested with `tests/minimal_test.ods` and
`tests/minimal_test.xlsx` to verify parity. See the
[Test Assets](tests/README.md) documentation for details.

### Test Patterns

- Use `include_bytes!` for unit tests to avoid I/O overhead.
- Use file paths for integration tests.
- Create both positive (expected behavior) and negative (error handling, edge
  cases) test cases.

## Code Style

Run these checks before submitting:

```bash
cargo fmt --check
cargo clippy --all-targets -- -D warnings
```

CI runs the checks in `.github/workflows/ci.yml` on every pull request. Fix
any red check before requesting review.

See [docs/coding-patterns.md](docs/coding-patterns.md) for architecture
conventions, naming patterns, and domain-specific guidelines.

## Documentation

Keep documentation in sync with the code — a change is not complete until the
docs match it:

| Change | Update |
| ------ | ------ |
| New/changed lint rule | `sheetlint/README.md` rule reference; `docs/coding-patterns.md` if it introduces a pattern |
| CLI args or behavior | the tool README (`sheetlint`/`sheetstats`/`sheetcli`); `README.md` if the getting-started experience changes |
| Config format (`sheetlint.toml`) | `docs/coding-patterns.md`; the relevant tool README |
| Module layout or new crate | `ARCHITECTURE.md`; `docs/README.md`; `README.md` project structure |
| Dependencies or commands | `README.md`; this guide |
| Engine behavior that differs between Excel and LibreOffice | `xlslang/docs/FIDELITY.md`; a probe in `xlslang/scripts/probe_dialect.py` |

## Changelog Fragments

Each PR adds one fragment to `changelog.d/`, named after its topic (e.g.
`changelog.d/ods-repeated-rows.md`). The fragment uses the
[Keep a Changelog](https://keepachangelog.com/) section headings (`Added`,
`Changed`, `Deprecated`, `Removed`, `Fixed`, `Security`), listing only the ones
the PR touches:

```markdown
### Added

- Add CALC203 rule for detecting double operators in formulas.

### Fixed

- Fix VUL604 false positive on XLOOKUP via word-boundary checking.
```

Before each release, fragments are merged into the `[Unreleased]` section of
`CHANGELOG.md` and deleted.

## Pull Request Process

1. Fork the repository and create a feature branch.
2. Make your changes following the coding patterns.
3. Ensure all tests pass with parity check.
4. Add a changelog fragment.
5. Submit a PR using the PR template.

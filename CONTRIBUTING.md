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

The Python bindings are tested with pytest against the built extension:

```bash
python -m venv .venv && . .venv/bin/activate
pip install maturin pytest
maturin develop -m sheetrs-py/Cargo.toml
pytest sheetrs-py
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
- A fixture for code that decides by looking for a name in text (a shortcut that skips a
  formula that does not mention a sheet) uses names that cannot occur in the text around them:
  `Zeta`, not `T` or `U`, which any function name contains.

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
| New/changed lint rule | `sheetrs-linter/README.md` rule reference; `docs/coding-patterns.md` if it introduces a pattern |
| CLI args or behavior | the tool README (`sheetlint`/`sheetstats`/`sheetedit`); `README.md` if the getting-started experience changes |
| Config format (`sheetlint.toml`) | `docs/coding-patterns.md`; the relevant tool README |
| Module layout or new crate | `ARCHITECTURE.md`; `docs/README.md`; `README.md` project structure |
| Dependencies or commands | `README.md`; this guide |
| API of the Python or WASM bindings | `sheetrs-py/README.md` and `sheetrs-py/python/sheetrs/_sheetrs.pyi`, or `sheetrs-wasm/README.md`; a pytest case for Python |
| Engine behavior that differs between Excel and LibreOffice | `sheetrs-compiler/docs/FIDELITY.md`; a probe in `sheetrs-compiler/scripts/probe_behavior.py` |

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

## Releases

1. Bump the version: `cargo set-version --workspace X.Y.Z` (cargo-edit; it also updates the
   `version` of the path dependencies), merge the changelog fragments into `CHANGELOG.md`.
2. Tag `vX.Y.Z` on `main` and push the tag. `.github/workflows/release.yml` checks that the tag
   matches the workspace version, builds the executables, Python wheels and sdist and the npm
   tarball, attaches them to a draft release and publishes it.
3. Publishing to the registries is manual, with the maintainer's credentials:
   - crates.io, in dependency order: `sheetrs-formula`, `sheetrs-core`, `sheetrs-stats`,
     `sheetrs-linter`, `sheetrs-edit`, `sheetrs-engine`, `sheetrs-compiler`
     (`sheetrs-wasm` and `sheetrs-py` have `publish = false`).
   - PyPI: `twine upload` the wheels and sdist of the release (or configure trusted publishing).
   - npm: `npm publish` the tarball of the release.

## Pull Request Process

1. Fork the repository and create a feature branch.
2. Make your changes following the coding patterns.
3. Ensure all tests pass with parity check.
4. Add a changelog fragment.
5. Submit a PR using the PR template.

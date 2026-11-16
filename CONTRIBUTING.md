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
# Run all tests (debug mode — safe when not all rules are enabled); the Python bindings
# are tested with pytest, below
cargo test --workspace --exclude sheetrs-py

# Run tests in release mode (recommended when all rules are enabled)
cargo test --release --workspace --exclude sheetrs-py
```

The Python bindings are tested with pytest against the built extension:

```bash
python -m venv .venv && . .venv/bin/activate
pip install maturin pytest
maturin develop -m sheetrs-py/Cargo.toml
pytest sheetrs-py
```

### Environment

`scripts/setup-dev-env.sh` prepares a Debian or Ubuntu machine: Rust 1.99 (the channel of
`rust-toolchain.toml`, as the default toolchain) and LibreOffice 26.8, the reference of the project,
with the checksum verified. `scripts/libreoffice-server.sh` starts the LibreOffice that the scripts of
the oracles and of the fixtures talk to (`sheetrs-io/tests/assets/oracle/make_oracle.py`,
`sheetrs-edit/tests/assets/replace/make_fixtures.py`, run with the Python of LibreOffice). The local machine and the cloud environment of Claude Code are both described in
[docs/environment.md](docs/environment.md).

### Parity and fidelity

Two requirements, and a rule about how to test them:

- **Parity of capabilities.** What SheetRS does for an XLSX it does for an ODS, and the other way
  round: a feature that only one format has is a bug.
- **Fidelity of behavior.** SheetRS gives the result **LibreOffice** gives (the reference version is
  in [ARCHITECTURE.md](ARCHITECTURE.md#reference-behavior)), whatever the algorithm or the
  data structure behind it. A difference exists only if the maintainer approved it, and then it is
  recorded in [docs/approved-differences.md](docs/approved-differences.md) and described where the
  behavior is documented. A difference nobody approved is a bug, even if it is convenient, hard to
  fix or unlikely to matter; one that cannot be fixed yet is reported to the maintainer for a
  decision, not shipped as if it were settled.
- **The reference is the workbook saved and opened again.** What LibreOffice holds in memory right after
  an edit (through its API or its screen) is not the reference: the same workbook, saved and opened by
  someone else, shows other values (a reference that went is `#REF!` in memory and `COUNT` of it is `#REF!`;
  opened again it is 0), so the memory is not reproducible and its behavior is a defect of LibreOffice, not
  something to copy. The expected result of an operation (a write, a replacement, a deletion of rows,
  columns, sheets or names, hiding or showing, a conversion of formulas to values, a calculation) is
  what a user sees when the workbook it leaves is **saved and opened again** in LibreOffice (`storeToURL`,
  then load, then read the values with the default load), in both formats. A hard recalculation of that file
  (`calculateAll`) is a cross-check; where it differs from the default load the default load is the
  expected result, and the case is reported. An oracle script that edits through UNO therefore saves,
  loads again and reads there; the values read from the document in memory are, at most, recorded as
  informational evidence, labeled so, and no test compares with them.
- **The calculation mode is LibreOffice's.** By default a run keeps the results a file caches for the formulas
  nothing reaches and calculates what an edit or an input reaches (what `LibreOffice` does when it opens a file and when a
  cell changes); `calculate_all` emulates its `calculateAll()` (every formula is calculated). The result a file caches is
  therefore not a goal of SheetRS but a behavior to emulate, in both modes, and where the default shows other values than
  `LibreOffice` after the same edit (`reread_default` of the oracle) the case is recorded in
  [docs/approved-differences.md](docs/approved-differences.md) as proposed, not approved.
- **A difference with LibreOffice is adjusted, not asked about** (a human developer and an agent alike).
  When SheetRS and LibreOffice differ and nothing prevents it, the difference is fixed in the same work. What
  prevents it is a technical impediment, a significant increase of complexity, or a regression or a
  degradation that has to be validated. Only then is a decision requested, and the request carries the
  justification for keeping the difference (permanent, or temporary, as when it is resolved in another pull
  request), says what is being asked, and says what the developer resolves without asking. With none of those
  reasons the developer resolves it.
- **The source code of LibreOffice is not read.** It is forbidden to a human developer and to an agent, in any
  version and in any form (its repository, a mirror, an excerpt, a tool that attaches it). The behavior is
  learned by running LibreOffice and from its own results, by a **clean room**: the behavior is observed with
  probes and oracles (`sheetrs-io/tests/assets/oracle/make_oracle.py` and the scripts of the fixtures) and written
  down as a specification of inputs and results, with no code and no internal name; the implementation is written
  from that specification and from the observations, never from the code; and every behavior is pinned by an
  oracle of results that LibreOffice itself gave. A developer who has seen the code of a function does not
  implement it. If a behavior cannot be derived from observations, it is reported to the maintainer as a pending
  decision, with what was tried: reading the code is not the way out.
- **How it is tested.** A file is tested against what LibreOffice reports for that file: the
  fixtures are made by LibreOffice and `sheetrs-io/tests/assets/oracle/make_oracle.py` records what it
  says of each cell (see `sheetrs-query/tests/libreoffice_oracle.rs`); the effect of an edit is tested against
  the file LibreOffice saves after making the same edit and what it shows when it opens it again
  (`sheetrs-edit/tests/assets/write_to/make_oracle.py`, see the rule above). A capability is tested on both
  formats, each with its own expectation.

Test assets and their description: [Test Assets](tests/README.md).

### Generators of fixtures are deterministic

A binary fixture (XLSX, ODS) that a script makes (`gen_*.py`, `make_*.py`: openpyxl builds it and
LibreOffice recalculates and saves it) is committed, because the tests have no LibreOffice at run
time and its cached values are the oracle. Running the script again must therefore leave `git status`
clean: a generator that writes the time of its run into the zip entries or the document properties
turns every regeneration into a diff of binary files that changed no cell.

- Every generator calls `fixture_determinism.normalize(path)` ([scripts/fixture_determinism.py](scripts/fixture_determinism.py),
  standard library only) on each workbook it leaves in the repository. It makes the date of every zip
  entry, `created`/`modified` of the document properties, the date fields LibreOffice writes into the
  headers and the order of the master pages of an ODS (taken from a hash table) constants. A date a
  fixture is about (`tests/lint_book.*` is old on purpose) goes in `DELIBERATE_MODIFIED`.
- The output is the same for the same LibreOffice build; the version of the build that wrote the file
  (`Application`, `meta:generator`) stays, so a new build shows up in the diff, as it should.
- Before committing a fixture: `scripts/fixture_determinism.py --check` lists any tracked workbook that
  still depends on the moment it was generated. Run a new generator twice and compare the bytes.

### Test Patterns

- Use `include_bytes!` for unit tests to avoid I/O overhead.
- Use file paths for integration tests.
- Create both positive (expected behavior) and negative (error handling, edge
  cases) test cases.
- **Name** a test for the behavior it checks, in a few words (`locale_flag_is_refused_at_run_time`), not
  for how the code reaches it. The name says **what the test proves, by itself**: it does not start with
  `test` (the module or the `tests/` directory already says it is a test), it has no comparative filler
  (`too`, `also`: say the behavior, not that it resembles another test) and it is not a long sentence that
  walks through the setup: short, concrete, stating the behavior (`a_bad_locale_tag_is_a_value_error`,
  `unreached_formulas_keep_their_result`). The same rule in Python: pytest collects `[a-z]*` functions in
  `sheetrs-py/tests` (`python_functions` in `pyproject.toml`), so a helper, a class or a constant starts
  with an underscore or a capital and is never collected. The Python tests prove the interface contract
  only (wiring, defaults, the mapping of errors and warnings); what the engine, the planner or the readers
  do is tested in Rust.
- **Test behavior, not design.** A test fails when what a user or a caller observes changes (a result, an
  error, an output). It does not pin that something does not exist or how it is built; a decision of
  design is recorded in the docs.
- **One behavior per test.** A test that checks several unrelated things is split.
- A fixture for code that decides by looking for a name in text (a shortcut that skips a
  formula that does not mention a sheet) uses names that cannot occur in the text around them:
  `Zeta`, not `T` or `U`, which any function name contains.

## Code Style

Run these checks before submitting:

```bash
cargo fmt --check
cargo clippy --workspace --exclude sheetrs-py --all-targets -- -D warnings
RUSTDOCFLAGS="-D warnings" cargo doc --workspace --exclude sheetrs-py --no-deps --document-private-items
```

Clippy runs with `clippy::all` and `clippy::pedantic` (`[workspace.lints.clippy]` in the root
`Cargo.toml`). The few pedantic lints that are allowed workspace-wide carry their reason next to
the setting; an `#[allow]` anywhere else needs a comment that says why the code is right and the
lint is not.

CI runs the checks in `.github/workflows/ci.yml` only when asked, so that a push does not build,
test or lint for nothing: put `[ci]` in the message of the last commit (a push) or in the title of
the pull request, add the label `ci` to the pull request, or start the workflow by hand (Actions >
CI > Run workflow). It runs the formatting, clippy and documentation checks, the tests of the
workspace (without `sheetrs-py`), the checks of the WebAssembly package and its tests in Node and
Chromium, the pytest suite of the Python binding and the build of its wheels (Linux x86_64 and
aarch64; macOS and Windows are out of scope for now, issue #99), and the determinism of the fixtures.
Fix any red check before requesting review.

See [docs/coding-patterns.md](docs/coding-patterns.md) for architecture
conventions, naming patterns, and domain-specific guidelines.

### API stability

The API is **experimental and unstable in every layer** (the command line tools, the Rust libraries,
the Python package and the WebAssembly package): names, parameters, defaults and results may change
without notice.

## Documentation

Keep documentation in sync with the code — a change is not complete until the
docs match it:

| Change | Update |
| ------ | ------ |
| New/changed lint rule | `sheetrs-linter/README.md` rule reference; `docs/coding-patterns.md` if it introduces a pattern (rules live in `sheetrs-linter/src/rules`) |
| CLI args or behavior | the tool README (`sheetrs lint`/`stats`/`query`/`edit`, `sheetc`); a subcommand lives in the `cli` module of its crate and `sheetrs-cli` only assembles them; `README.md` if the getting-started experience changes |
| Config format (`sheetrs.toml`) | `docs/coding-patterns.md`; the relevant tool README |
| Module layout or new crate | `ARCHITECTURE.md` (a module goes in `sheetrs-io` only if two tools need it); `docs/README.md`; `README.md` project structure |
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

What CI does and what stays manual:

| Step | Who | How |
| ---- | --- | --- |
| Version | you | `cargo set-version --workspace X.Y.Z` (cargo-edit; it also updates the `version` of the path dependencies). It is the only place: the Python wheel and the npm package take the version of the workspace. |
| Changelog | you | merge the fragments of `changelog.d/` into a `## [X.Y.Z]` section of `CHANGELOG.md` and delete them: that section is the first part of the release notes (GitHub adds the generated list of pull requests) |
| Check the path | optional | Actions > Release > Run workflow on the branch is a dry run (builds and smoke-tests every artifact, creates no release). The jobs of `ci.yml` already build the executables, the wheels and the npm package of every PR, and the release created by a tag stays a draft until its smoke tests pass, so run it only when the change touches `release.yml` itself |
| Tag | CI | when the workspace version reaches `main` and has no `vX.Y.Z` tag, the job `plan` of `release.yml` creates the tag on that commit; a tag pushed by hand also works, and must match the workspace version |
| Build and release | CI (`.github/workflows/release.yml`) | the tag must match the workspace version; builds for Linux x86_64 and aarch64 the executables (`sheetrs`, `sheetc`, `sheetrs-rt` and the two base modules of the WebAssembly target, with `.sha256`), the `abi3` wheels (tested in clean environments of Python 3.13, 3.14 and 3.15), the sdist and the npm tarball; attaches them to a draft release, smoke-tests them and publishes the release. If a job fails, the release stays a draft: delete the draft and the tag, fix, and re-run the workflow of the push to `main` (it creates the tag again) |
| crates.io, PyPI, npm | you, until [#108](https://github.com/cosmoscalibur/sheetrs/issues/108) | needs the credentials of the maintainer: crates.io in dependency order (`sheetrs-formula`, `sheetrs-io`, `sheetrs-protocol`, `sheetrs-engine`, `sheetrs-query`, `sheetrs-stats`, `sheetrs-linter`, `sheetrs-answer`, `sheetrs-compiler`, `sheetrs-edit`; `sheetrs-wasm`, `sheetrs-rt-wasm` and `sheetrs-py` have `publish = false`), PyPI with `twine upload` of the wheels and sdist of the release (or trusted publishing), npm with `npm publish` of the tarball |

macOS and Windows artifacts are out of scope for now: [#99](https://github.com/cosmoscalibur/sheetrs/issues/99).

## Pull Request Process

1. Fork the repository and create a feature branch.
2. Make your changes following the coding patterns.
3. Ensure all tests pass, and that what changed was checked against LibreOffice on both formats.
4. Add a changelog fragment.
5. Submit a PR using the PR template.

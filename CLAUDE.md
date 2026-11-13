# CLAUDE.md — SheetRS Suite

Claude Code entry file (hot memory) carrying **SheetRS-specific context only**.
General methodology — planning, implementation, review, commit, PR — comes from
the global skills and configuration. Project constraints,
conventions, and the contribution process live in the docs referenced below;
this file only points to them.

## Load before working

Read these before planning or implementing any change:

- `docs/coding-patterns.md` — SoC, WalkerRule pattern, formula normalization,
  rule config parameters, shared helpers, CLI/library conventions.
- `CONTRIBUTING.md` — testing (capability parity, fidelity to LibreOffice, TDD), documentation maintenance,
  changelog fragments, PR process.

Load on demand:

- `README.md` — tech stack, prerequisites, build/run/test/lint commands.
- `ARCHITECTURE.md` — workspace layout, modules, design principles.
- `docs/README.md` — documentation index (profiling, benchmarks, test assets,
  WASM demo).

## Benchmarks

- Benchmarks run on generated workbooks, so any checkout reproduces them:
  `scripts/bench.py fixtures DIR` then `scripts/bench.py run DIR`.
- Read `docs/performance.md` before benchmarking. A benchmarked change updates
  the results table of that document, and `sheetrs-linter/docs/PROFILING.md` when it targets the
  cost of a rule or phase.

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
- `CONTRIBUTING.md` — testing (ODS/XLSX parity, TDD), documentation maintenance,
  changelog fragments, PR process.

Load on demand:

- `README.md` — tech stack, prerequisites, build/run/test/lint commands.
- `ARCHITECTURE.md` — workspace layout, modules, design principles.
- `docs/README.md` — documentation index (profiling, benchmarks, test assets,
  WASM demo).

## Benchmarks

- `tests/production/` is git-ignored, so a git worktree lacks it. Symlink it
  from the main checkout before benchmarking:
  `ln -s <main-checkout>/tests/production tests/production`.
- Read `scripts/README.md` before benchmarking. A benchmarked change updates
  its Comparisons entry, and `sheetlint/docs/PROFILING.md` when it targets the
  cost of a rule or phase.

## Agent readiness

Re-run the `agent-readiness` skill when agent-facing tooling is added or
removed (CI/workflows, a linter/formatter/type-checker, test infrastructure,
documentation structure, build/devcontainer setup, tooling dependencies).
Persist reports under `docs/agent-readiness/`.

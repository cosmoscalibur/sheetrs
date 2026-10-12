# CLAUDE.md — SheetRS Suite

Claude Code entry file (hot memory) carrying **SheetRS-specific context only**.
General methodology — planning, implementation, review, commit, PR — comes from
the `tributi-hardness` skills and global configuration. Project constraints,
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

## Agent readiness

Re-run `agent-harness:agent-readiness` when agent-facing tooling is added or
removed (CI/workflows, a linter/formatter/type-checker, test infrastructure,
documentation structure, build/devcontainer setup, tooling dependencies).
Persist reports under `docs/agent-readiness/`.

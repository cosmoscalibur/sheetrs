# Documentation Index

This directory serves as the knowledge base for the SheetRS Suite. For
getting-started instructions, see the root [README.md](../README.md).

## Architecture & Design

- [Architecture](../ARCHITECTURE.md) — workspace structure, module layout,
  design principles (format-agnostic rules, streaming readers).
- [Coding Patterns](coding-patterns.md) — conventions, naming, SoC principle,
  WalkerRule trait pattern, CLI and library guidelines.

- [Spike: interpreter, JIT and LLVM](spike-execution-backends.md) — what a direct mode, a JIT and an
  LLVM compiler would add to the engine, with the measurements behind the verdicts.

## API

- [API alignment](api-alignment.md) — the rules that keep the command line, Rust, Python and WebAssembly
  aligned, and what each operation is called in every layer. The API is experimental and unstable.

## Design references

- [Reference projects](references.md) — projects that inspire the features and the shape of the API.

## Environment

- [Development environment](environment.md) — the local machine and the cloud environment of
  Claude Code: Rust, LibreOffice 26.8, network, disk, profiling.

## Compiler and engine (sheetc)

- [sheetrs-compiler README](../sheetrs-compiler/README.md) — build, run, inputs/outputs, options.
- [Design](../sheetrs-compiler/docs/DESIGN.md) — pipeline, IR, hybrid compile/run time,
  ordering, reader contract, roadmap.
- [Fidelity](../sheetrs-compiler/docs/FIDELITY.md) — every difference between Excel and
  LibreOffice the engine models, with the evidence behind it.

## Recipes

- [Recipes](recipes.md) — recalculate before linting, the answer of a model as a workbook or as
  JSON, find and replace.

## WebAssembly

- [sheetrs-wasm README](../sheetrs-wasm/README.md) — every tool in the browser and in Node: the
  `Book`, its options, what is not there; `Book.compile` and the exports of the module of a
  workbook.
- [sheetrs-compiler README, WebAssembly target](../sheetrs-compiler/README.md#webassembly-target) —
  `sheetc build --target wasm`.

## Reading and finding cells

- [sheetrs-query README](../sheetrs-query/README.md) — `sheetrs query`: ranges, sheets, properties, lists of sheets and names,
  display, find.
- [sheetrs-protocol README](../sheetrs-protocol/README.md) — the typed JSON contract.

## Analysis & Profiling

- [Comparative Report](../sheetrs-linter/comparative_report.md) — rule comparison
  with PerfectXL Risk Finder and ExceLint.
- [Profiling](../sheetrs-linter/docs/PROFILING.md) — phase breakdown, per-rule
  sample share, single-pass walker architecture.
- [Performance](performance.md) — the current time and memory of every tool (linter, statistics, query, editor,
  compiler and engine, Python, WebAssembly) on generated workbooks, the same flows in LibreOffice, and the known
  limits; [scripts/README.md](../scripts/README.md) lists the scripts that measure it.

## Testing

- [Test Assets](../tests/README.md) — test file descriptions, usage patterns,
  known limitations.

## Bindings

- [sheetrs-py](../sheetrs-py/README.md) — Python package: `Book` (`range`, `find`, `lint`, `stats`, and
  the operations that accumulate until `save`: `delete_*`, `formula_to_value`, `replace`, `calculate`).
- [sheetrs-wasm](../sheetrs-wasm/README.md) — WebAssembly module and npm package.

## WASM Demo

- [SheetRS WASM Demo](index.html) — browser-based demo using the sheetrs-wasm
  package: lint (report, table, SARIF), statistics, query, edit, calculate, evaluate and compile a
  workbook to a WebAssembly module and run it in the page. Build the package it reads with
  `scripts/package-wasm.sh --demo` (it goes to `docs/pkg/`) and serve this folder.

# Documentation Index

This directory serves as the knowledge base for the SheetRS Suite. For
getting-started instructions, see the root [README.md](../README.md).

## Architecture & Design

- [Architecture](../ARCHITECTURE.md) — workspace structure, module layout,
  design principles (format-agnostic rules, streaming readers).
- [Coding Patterns](coding-patterns.md) — conventions, naming, SoC principle,
  WalkerRule trait pattern, CLI and library guidelines.

## Compiler and engine (sheetc)

- [sheetrs-compiler README](../sheetrs-compiler/README.md) — build, run, inputs/outputs, options.
- [Design](../sheetrs-compiler/docs/DESIGN.md) — pipeline, IR, hybrid compile/run time,
  ordering, reader contract, roadmap.
- [Fidelity](../sheetrs-compiler/docs/FIDELITY.md) — every difference between Excel and
  LibreOffice the engine models, with the evidence behind it.

## Analysis & Profiling

- [Comparative Report](../sheetrs-linter/comparative_report.md) — rule comparison
  with PerfectXL Risk Finder and ExceLint.
- [Profiling](../sheetrs-linter/docs/PROFILING.md) — phase breakdown, per-rule
  sample share, single-pass walker architecture.
- [Benchmarks](../scripts/README.md) — benchmark scripts and historical
  performance data.

## Testing

- [Test Assets](../tests/README.md) — test file descriptions, usage patterns,
  known limitations.

## WASM Demo

- [SheetRS WASM Demo](index.html) — browser-based demo using the sheetrs-wasm
  crate. Supports lint and stats tools on XLSX/ODS files.

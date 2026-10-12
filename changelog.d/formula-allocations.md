### Added

- Add `LinterContext::formula_upper`, which converts each cell's formula to uppercase once for all rules. `LinterContext` gains a private field, so it is built through `Default`.
- Add a `profiling` Cargo profile (release with fat LTO and one codegen unit) for benchmarks and profiling.

### Changed

- Reduce per-cell allocations in formula rules (VUL602, VUL603, VUL604, VUL606, CALC204, CALC205, REF307, INT401–403, CPX504) and in XLSX shared-formula translation: shared uppercase formula, precomputed function patterns, and regex replacements written straight into the output buffer. Production benchmark: about 5% fewer user-space instructions, same peak RAM.
- `scripts/benchmark.sh` builds with the `profiling` profile, runs each production file twice, and reports user-space instructions through `perf stat` (now required).

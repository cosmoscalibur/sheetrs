# Sheetlint Profiling & Benchmarking

This directory contains scripts for measuring the performance of `sheetlint`.

## Scripts

### `benchmark.sh`

The main benchmark script. Uses `/usr/bin/time -v` and `perf stat` for
comprehensive metrics:

- User time, System time, Elapsed time
- CPU usage %
- Peak RAM (Maximum resident set size)
- User-space instructions (`perf stat`), stable across sessions where time and
  RAM shift
- Binary size

The script builds `sheetlint` with the `profiling` Cargo profile (release with
fat LTO and one codegen unit), so binary size and timings refer to that build.
Each production file runs twice, with both runs' metrics reported; the
violations report comes from the first run.

**Requirements:** `time` package (provides `/usr/bin/time`) and `perf`

**Usage:**

```bash
./scripts/benchmark.sh
```

### Interpreting results

Compare the metrics against the latest entry in [Comparisons](#comparisons).
When the two runs of one build disagree, or a metric moves against the latest
entry without an explanation in the change, investigate before accepting the
result: debug the suspected code path, or profile it (requires `perf`):

```bash
cargo flamegraph --profile profiling -p sheetlint -- <file>
```

[PROFILING.md](../sheetlint/docs/PROFILING.md) holds the phase and per-rule
profile and the commands that produce it.

Time and RAM shift between sessions on the same machine: the same binary has
measured 1943 MB and 1990 MB peak RAM, and 37.5 s and 42.2 s user time, on the
XLSX file. When the latest entry was measured under different conditions,
rebuild it at its commit in the same session as the new measurement.

A change gets a new Comparisons entry only when it moves a metric.

## Production Output

`sheetlint` reports millions of violations for the production files, so its
output runs to hundreds of MB, like the reference outputs in
`tests/production/baselines/`. `benchmark.sh` writes that full output into its
`scripts/benchmark_*.txt` results file.

To compare outputs between two builds, use streaming tools (`sort -T <dir>`,
`diff`, `md5sum`) on the text output. Tools that load the whole document into
memory, such as `jq` on the `--format json` output, can exhaust RAM.

Two runs of the same build can differ: violation order varies (#33), so sort
both outputs before diffing, and REF310 may report a different set of
violations (#34), so a diff limited to REF310 lines and the summary counts is
expected.

## Comparisons

Benchmarks use production files in `tests/production/`. Each entry measures
the state after its change, so the previous entry is its baseline. Entries
from #30 on come from `benchmark.sh` with the `profiling` profile, mean of 2
runs, measured in one session; older entries are release builds.

### Per-cell formula allocations (#31)

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 39.5s | 38.3s |
| Peak RAM | 1943 MB | 1958 MB |
| User instructions | 364.7 G | 347.1 G |
| Binary size | 9.0M | - |

### Rust 1.99 lints and cleanups (#30)

Rebuilt at 9750d31 with the current `benchmark.sh`.

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 39.9s | 38.5s |
| Peak RAM | 1942 MB | 1958 MB |
| User instructions | 385.6 G | 367.2 G |
| Binary size | 9.0M | - |

### Arc+Box+dedup+shrink+ods-no-clone (4172267)

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 29.9s | 26.7s |
| Peak RAM | 1315 MB | 1329 MB |
| Binary size | 8.4M | - |

### ViolationData Refactoring (2cbc5fc)

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 27.2s | 23.3s |
| Peak RAM | 1528 MB | 1631 MB |
| Binary size | 8.4M | - |

### CALC201 Walker Migration (e9c34b9)

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 35.7s | 28.8s |
| Peak RAM | 1457 MB | 1594 MB |
| Binary size | 8.4M | - |

### Walker Integration (56e85b5)

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 30.4s | 25.2s |
| Peak RAM | 1744 MB | - |
| Binary size | 8.3M | - |

### Memory dependencies (6db5c17)

- User time: 55s
- Peak RAM: 1927MB
- Binary size: 8.6 MB

### Previous

- User time: 58.25s
- Peak RAM: 3616MB
- Binary size: 8.9 MB

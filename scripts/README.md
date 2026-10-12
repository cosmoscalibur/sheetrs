# Sheetlint Profiling & Benchmarking

This directory contains scripts for measuring the performance of `sheetlint`.

## Scripts

### `benchmark.sh`

The main benchmark script. Uses `/usr/bin/time -v` for comprehensive metrics:

- User time, System time, Elapsed time
- CPU usage %
- Peak RAM (Maximum resident set size)
- Binary size

**Requirements:** `time` package (provides `/usr/bin/time`)

**Usage:**

```bash
./scripts/benchmark.sh
```

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

Benchmarks use production files in `tests/production/`.

### Rust 1.99 lints and cleanups (#30)

Mean of 3 runs.

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 37.8s | 36.0s |
| Peak RAM | 1944 MB | 1959 MB |
| Binary size | 12.2M | - |

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

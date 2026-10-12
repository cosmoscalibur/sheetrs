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
- MD5 of each run's output
- Binary size

The script builds `sheetlint` with the `profiling` Cargo profile (release with
fat LTO and one codegen unit), so binary size and timings refer to that build.
Each production file runs twice, with both runs' metrics reported; the
violations report comes from the first run.

`sheetlint` runs with `scripts/benchmark.toml`, so changes to the root
`sheetlint.toml` example leave the measured rule work unchanged. A change to
its thresholds changes the output and the metrics, so the next Comparisons
entry states it.

**Requirements:** `time` package (provides `/usr/bin/time`) and `perf`

**Usage:**

```bash
./scripts/benchmark.sh                     # current tree only
./scripts/benchmark.sh --baseline develop  # also time a git ref, same session
```

`--baseline <ref>` builds the ref in a temporary git worktree (target dir
`target/baseline/`) and times it before the current tree. Its results carry
metrics only, without the violations report.

### Interpreting results

Compare the metrics against the latest entry in [Comparisons](#comparisons).
When the two runs of one build disagree, or a metric moves against the latest
entry without an explanation in the change, investigate before accepting the
result: debug the suspected code path, or profile it (requires `perf`):

```bash
cargo flamegraph --profile profiling -p sheetlint -- <file>
```

[PROFILING.md](../sheetlint/docs/PROFILING.md) holds the phase and per-rule
profile and the commands that produce it. A change that targets the cost of a
rule or phase updates that profile along with its Comparisons entry.

Time and RAM shift between sessions on the same machine: the same binary has
measured 1943 MB and 1990 MB peak RAM, and 37.5 s and 42.2 s user time, on the
XLSX file. User instructions shift less: the same binary has measured 339.3 G
and 335.8 G on the XLSX file, and 335.8 G and 333.0 G on the ODS file. When a
metric differs from the latest entry by no more than its shift, run with
`--baseline` to tell the change from the session.

[Comparisons](#comparisons) is incremental: a new entry measures only the
current tree. A change gets an entry when a metric moves beyond the spread
between its two runs; a change within that spread is noise and gets no entry.

## Production Output

`sheetlint` reports millions of violations for the production files, so its
output runs to hundreds of MB, like the reference outputs in
`tests/production/baselines/`. `benchmark.sh` writes that full output into its
`scripts/benchmark_*.txt` results file.

Two runs of one build produce byte-identical output. Each
[Comparisons](#comparisons) entry records the output MD5 that `benchmark.sh`
reports, so a change checks its output by comparing MD5 against the latest
entry. When the MD5 differs and the change does not explain it, regenerate both
outputs and compare them with streaming tools (`diff`) to find the difference.
Tools that load the whole document into memory, such as `jq` on the
`--format json` output, can exhaust RAM.

## Comparisons

Benchmarks use production files in `tests/production/`. Each entry measures
the state after its change, so the previous entry is its baseline. Entries
from #30 on come from `benchmark.sh` with the `profiling` profile, mean of 2
runs, measured in one session; older entries are release builds.

### XML entity references in element text (#48)

XLSX formulas and text keep their entity references and the whitespace next
to them, so CALC201, CPX506, CPX508, CPX509 and VUL603 report the same counts
as ODS, and DATA705 reports 40,614 XLSX violations (none before). ODS output
changes in DATA704, DATA705 and EXT802. Peak RAM on XLSX grows with the larger
report. The previous entry's build, run once in the same session, measured
335.8 G (XLSX) and 333.0 G (ODS) user instructions and 1799 MB (XLSX) peak RAM.

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 30.7s | 30.4s |
| Peak RAM | 1813 MB | 1821 MB |
| User instructions | 336.0 G | 332.9 G |
| Output MD5 | `27417019c56695b4857551f630a8ffb1` | `c51c9d7817650f251ff10455b3c548db` |
| Binary size | 9.2M | - |

### REF301 single formula scan (#37)

Output is byte-identical to 4fe45d7 (previous entry), with all rules and
with only REF301 enabled (50 XLSX and 38 ODS REF301 violations).

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 31.8s | 31.9s |
| Peak RAM | 1801 MB | 1819 MB |
| User instructions | 339.3 G | 335.8 G |
| Output MD5 | `c2e8390c660ac4627ed2a2225888963d` | `9d0adac635fbfb95cc61ee8158fdd1cf` |
| Binary size | 9.1M | - |

### quick-xml 0.42 and regex 1.13 (#44)

Output is identical to the previous entry. One further run of each build in
the same session, outside `benchmark.sh`, took 38.4 s (previous entry) and
38.6 s (this entry) user time on XLSX and 37.6 s and 37.4 s on ODS, with peak
RAM within 8 MB of each other.

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 40.8s | 37.9s |
| Peak RAM | 1945 MB | 1963 MB |
| User instructions | 354.8 G | 348.2 G |
| Binary size | 9.1M | - |

### Deterministic output and REF310 boundary matching (#33, #34)

REF310 reports 118,501 (XLSX) and 119,735 (ODS) violations; 82f57f2 reports
102,182 on the XLSX file.

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 40.1s | 38.3s |
| Peak RAM | 1949 MB | 1969 MB |
| User instructions | 362.1 G | 344.9 G |
| Binary size | 9.1M | - |

### Lock-free walker rules (#32)

Baseline rebuilt at 634d52e in the same session: 365.1 G (XLSX) and 347.3 G
(ODS) user instructions.

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 40.5s | 38.3s |
| Peak RAM | 1942 MB | 1957 MB |
| User instructions | 364.0 G | 346.4 G |
| Binary size | 9.0M | - |

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

# Sheetlint Profiling & Benchmarking

This directory contains scripts for measuring the performance of `sheetlint`
and `sheetedit`.

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

### `benchmark_edit.sh`

Times `sheetedit` modifications with the same metrics as `benchmark.sh`, built
with the `profiling` profile. Each scenario of its `SCENARIOS` list runs twice
on each production file: removing a sheet, removing a named range, converting
a whole sheet, converting columns and rows, and all of them in one command.

`sheetedit` modifies a copy of each production file in a temporary directory
and writes its output there; the script fails if a production file changes.
ZIP entries carry their write time, so the reported output MD5 hashes the
extracted parts rather than the archive.

**Usage:**

```bash
./scripts/benchmark_edit.sh
```

Its results go to the [Edit Comparisons](#edit-comparisons), apart from the
`sheetlint` ones.

### Interpreting results

Compare the metrics against the latest entry in [Comparisons](#comparisons).
When the two runs of one build disagree, or a metric moves against the latest
entry without an explanation in the change, investigate before accepting the
result: debug the suspected code path, or profile it (requires `perf`):

```bash
cargo flamegraph --profile profiling -p sheetrs-linter -- <file>
```

[PROFILING.md](../sheetrs-linter/docs/PROFILING.md) holds the phase and per-rule
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

## Edit Comparisons

`benchmark_edit.sh` entries, mean of 2 runs, measured in one session. The
scenario names are those of its `SCENARIOS` list.

### `sheetstats` counts while it reads (#73)

`sheetstats` hands the cells to a counter as the reader finds them instead of building the
workbook, and reports the end, comments, charts and tables of each sheet. Same counts as the
version that kept the cells, on the production files (user time, peak RAM):

| Version | XLSX | ODS |
|---------|------|-----|
| Cells kept in memory | 2.27s, 462 MB | 3.54s, 440 MB |
| Cells counted while reading | 2.15s, 22 MB | 3.47s, 5 MB |

### Rows and columns deleted, references adjusted (#73)

The deletion rewrites every sheet, because a formula of any sheet can read what goes; removing a
sheet does the same now that the references to it become `#REF!`. An XLSX sheet that names none of
the sheets edited and holds no reference across sheets (`Sheet1:Sheet3!A1`, which spans sheets it
does not name) is copied as it is after a scan of its text. `delete-rows` and `delete-columns`
cut `INPUTS`, `delete-rows-columns` adds `--recalc-on-open`, and `structure-plan` removes a sheet
and a name too. Before this entry `remove-sheets` took 0.00s on the XLSX and 2.38s on the ODS,
with 5 MB of RAM in both.

| Scenario | XLSX user time | XLSX peak RAM | XLSX user instructions | ODS user time | ODS peak RAM | ODS user instructions |
|----------|----------------|---------------|------------------------|---------------|--------------|-----------------------|
| `remove-sheets` | 1.17s | 6 MB | 22.3 G | 4.48s | 8 MB | 73.2 G |
| `remove-ranges` | 0.01s | 5 MB | 0.0 G | 2.27s | 5 MB | 33.9 G |
| `convert-sheet` | 3.33s | 8 MB | 47.3 G | 3.33s | 6 MB | 55.0 G |
| `convert-columns-rows` | 3.62s | 8 MB | 49.7 G | 4.00s | 5 MB | 63.3 G |
| `plan` | 3.43s | 9 MB | 50.0 G | 3.95s | 8 MB | 64.9 G |
| `delete-rows` | 7.58s | 9 MB | 110.5 G | 6.88s | 8 MB | 113.5 G |
| `delete-columns` | 7.54s | 9 MB | 108.2 G | 6.87s | 8 MB | 113.4 G |
| `delete-rows-columns` | 7.75s | 9 MB | 112.0 G | 7.30s | 8 MB | 119.6 G |
| `structure-plan` | 7.73s | 9 MB | 113.4 G | 7.07s | 8 MB | 118.1 G |

On the ODS the sheet names of `content.xml` are read in a second pass only when a 3D reference
needs them (it spans the sheets between its ends, which can come after the formula), so a
document without them pays nothing for it.

The peak memory of every scenario is about 9 MB on these production files: the sheets stream, one
row (ODS) or one cell (XLSX) at a time. A deletion costs about twice the time of `convert-sheet`.

Profile of `delete-rows` on the XLSX (`perf`, share of the samples of the run):

| Share | Where |
|-------|-------|
| 25% | compression of the output ZIP (`deflate`) |
| 21% | reading and rewriting the formulas (`SheetStructure::rewrite`, of which the parser 11%) |
| 15% | rebuilding each cell element to renumber it (`relocated_cell`) |
| 10% | reading the XML (`quick-xml`) |
| 7% | translating the cells of shared formulas (`shared_text`) |

Plan to bring the time down, in the order of its return:

1. Compress the output at a fast level (or store it): the first row of the profile.
2. Rewrite each distinct formula text once (workbooks repeat patterns) instead of parsing it for
   every cell.
3. Change only the `r` attribute of a cell in place, as text, instead of rebuilding its attributes.
4. ODS: copy a row that holds no formula as it is, without buffering it, and skip the formulas
   that name none of the sheets edited, as the XLSX does for its sheets.
5. XLSX: rewrite the sheets in parallel (they are independent parts), at the price of a sheet in
   memory per thread.

### Formula conversion and ODS modifications (#59)

Before streaming, a release build of this change peaked at 366 MB of RAM
converting every XLSX sheet, and at 1.0 GB to 1.2 GB on every ODS
modification.

| Scenario | XLSX user time | XLSX peak RAM | ODS user time | ODS peak RAM |
|----------|----------------|---------------|---------------|--------------|
| `remove-sheets` | 0.00s | 5 MB | 2.38s | 5 MB |
| `remove-ranges` | 0.00s | 4 MB | 2.41s | 5 MB |
| `convert-sheet` | 2.95s | 8 MB | 3.43s | 5 MB |
| `convert-columns-rows` | 2.95s | 6 MB | 4.01s | 5 MB |
| `plan` | 2.94s | 8 MB | 3.41s | 5 MB |

| Scenario | XLSX user instructions | ODS user instructions |
|----------|------------------------|-----------------------|
| `remove-sheets` | 0.02 G | 34.0 G |
| `remove-ranges` | 0.02 G | 34.3 G |
| `convert-sheet` | 40.9 G | 54.7 G |
| `convert-columns-rows` | 40.5 G | 63.0 G |
| `plan` | 40.9 G | 54.4 G |

| Scenario | XLSX output MD5 (parts) | ODS output MD5 (parts) |
|----------|-------------------------|------------------------|
| `remove-sheets` | `53d901757d6d602ee15d7166be91ce60` | `8ae4b206148c3534b4259867ce24ac52` |
| `remove-ranges` | `ff21a6bf6b21667823c49410e5d63eb4` | `ef43d9d784bb02504ab2ce0264ca00c5` |
| `convert-sheet` | `d199f265d201d5532f58c6f89247631a` | `64fab675d59578eed46af1d3dec01475` |
| `convert-columns-rows` | `d169b5c383777e92e3de1592114b83dc` | `baab6a8d29f11853167551e2167aa803` |
| `plan` | `489dacde8c6b2adda89d05f3a05527f0` | `5161ae6590b80e127609177939398a8e` |

Binary size: 3.4M.

## Comparisons

Benchmarks use production files in `tests/production/`. Each entry measures
the state after its change, so the previous entry is its baseline. Entries
from #30 on come from `benchmark.sh` with the `profiling` profile, mean of 2
runs, measured in one session; older entries are release builds.

### Output MD5 of the linter since `ac80165`

The linter's output MD5 differs from the entries below on purpose: the reader reads the cached
string result of a formula as text, a number in a Text-formatted cell as a number, and CALC202
takes the references of a formula from its AST. What each rule reports, and why, is in the
documentation of its module (`cargo doc -p sheetrs-core --open`, `sheetrs_core::rules`).

### Compiler and engine crates, rule corrections (#73)

The output MD5 differs from the previous entry on purpose (see *Output MD5 of the linter since `ac80165`*). The shared
formulas of an XLSX are translated by a scanner instead of a regular expression, and the ODS
reader stops looking for the null date at the first table. Against the previous entry (#39),
user instructions are 0.5% lower on XLSX and 3% higher on ODS, and peak RAM falls by a third.

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 25.8s | 26.5s |
| Peak RAM | 1239 MB | 1215 MB |
| User instructions | 321.3 G | 332.6 G |
| Output MD5 | `b4f753355076ca606e788aa446f28223` | `afb440b8a13a69cec011b57ce37b7517` |
| Binary size | 9.2M | - |

The ODS MD5 was `3dd17a729a43f97032e27cd55e357a82` before the ODS reader read a range that repeats
its sheet (`[$S.A1:$S.B1]`) as the one that leaves it out (`[$S.A1:.B1]`). The file holds
`SUM(INPUTS!DA36474:DE36474)` written both ways, so VUL601 reports 14 more duplicated formulas;
nothing else changes there. CPX509 then counts the characters of a formula without the namespace
prefix of its functions (`_xlfn.`, `COM.MICROSOFT.`), as the application shows them: the MD5s
before it were `3d3b557c191acff9c5bc5a7e4eb7c223` (XLSX) and `5280d61581b3db4464737826ce3a9b1a`
(ODS), and only CPX509 messages differ (both formats report the same 26 formulas now).

### Violation sink for walker hooks (#39)

`WalkerRule` hooks push into the walker's `Vec<Violation>` instead of returning
one. The output MD5 matches the previous entry on both formats. The previous
entry's build, run in the same session on XLSX, measured 328.8 G user
instructions, 32.5s user time and 1800 MB peak RAM.

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 32.5s | 32.9s |
| Peak RAM | 1805 MB | 1806 MB |
| User instructions | 323.0 G | 321.9 G |
| Output MD5 | `c50fafb3ffda7c95de2ed74cf9dc32fc` | `de1b6bd593326feb06f7566a970406ac` |
| Binary size | 9.1M | - |

### CPX504–CPX509 cell scope and CPX509 character length (#67)

CPX504–CPX509 report with cell scope, so the text output lists them under
their formula cell in row and column order and their messages drop the
`at <cell>` suffix. CPX509 measures characters instead of UTF-8 bytes: 187
(XLSX) and 186 (ODS) of its messages report a shorter length, and five
formulas of 256–257 bytes (`INPUTS!AQ2909`, `AN2331`, `AN2223`, `AO21998`,
`AN2281`) no longer trigger it in either format. Every other violation matches
the previous entry's output, compared by sheet, scope, cell, rule and message.
The previous entry's build, run once in the same session, measured 331.8 G user
instructions and 1817 MB peak RAM on XLSX.

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 33.4s | 34.3s |
| Peak RAM | 1802 MB | 1802 MB |
| User instructions | 328.8 G | 327.4 G |
| Output MD5 | `c50fafb3ffda7c95de2ed74cf9dc32fc` | `de1b6bd593326feb06f7566a970406ac` |
| Binary size | 9.1M | - |

### Providers run for enabled dependents (#53)

Output is identical to the previous entry. The previous entry's build, run in
the same session, measured 331.8 G (XLSX) and 330.0 G (ODS) user instructions.
A disabled provider now runs whenever a rule depending on it is enabled. Single
XLSX runs outside `benchmark.sh`, this entry against the previous one:

| Configuration | User time | Peak RAM |
|---------------|-----------|----------|
| `disabled_rules = ["CALC202"]` | 37.6s / 21.6s | 1816 MB / 1580 MB |
| `enabled_rules` = ERR103, REF310, DATA706, CPX505/506/508/509 | 15.5s / 4.6s | 1186 MB / 637 MB |

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 33.5s | 33.9s |
| Peak RAM | 1814 MB | 1821 MB |
| User instructions | 331.8 G | 330.0 G |
| Output MD5 | `27417019c56695b4857551f630a8ffb1` | `c51c9d7817650f251ff10455b3c548db` |
| Binary size | 9.2M | - |

### Allocation-free column letters (#40)

Output is identical to the previous entry. The previous entry's build, run
once in the same session, measured 336.0 G (XLSX) and 332.9 G (ODS) user
instructions and 32.5 s (XLSX) user time; one interleaved XLSX run of each
build measured 160.6 G (previous entry) and 158.1 G (this entry) user cycles.
User time ran under external load in this session.

| Metric | XLSX | ODS |
|--------|------|-----|
| User time | 34.3s | 33.6s |
| Peak RAM | 1813 MB | 1817 MB |
| User instructions | 331.8 G | 330.0 G |
| Output MD5 | `27417019c56695b4857551f630a8ffb1` | `c51c9d7817650f251ff10455b3c548db` |
| Binary size | 9.2M | - |

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

Output is byte-identical to 126d3c1 (previous entry), with all rules and
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

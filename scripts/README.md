# Scripts

Scripts to prepare a machine, to build the packages and to measure performance. The numbers, the method, the
machine and the contrast with LibreOffice are in [docs/performance.md](../docs/performance.md); this file lists
what each script does.

| Script | What it does |
| ------ | ------------ |
| `bench.py` | The benchmarks: `fixtures DIR` generates the workbooks (nothing private is needed), `run DIR [groups]` measures `sheetrs lint`, `stats`, `query`, `edit`, `sheetc calculate` and `build`, the Python `Book`, the WebAssembly package and the LibreOffice flows with dist-profile binaries, and prints Markdown tables (wall time and peak memory, min-max of the runs) |
| `measure_libreoffice_flow.py` | The same flows in LibreOffice headless through UNO (times per phase, peak RSS) next to `sheetrs` and `sheetc`; run it with the Python of LibreOffice (`/opt/libreoffice26.8/program/python`) on a workbook of `sheetrs-compiler/scripts/gen_large_fixture.py` |
| `benchmark.toml` | The lint configuration of the benchmarks: thresholds fixed, so every run does the same rule work |
| `setup-dev-env.sh` | Prepares a Debian or Ubuntu machine: the Rust toolchain of `rust-toolchain.toml`, LibreOffice 26.8, `perf` and `valgrind` ([docs/environment.md](../docs/environment.md)) |
| `setup-claude-env.sh` | The setup script of the Claude Code environment (pasted whole into the environment settings) |
| `libreoffice-server.sh` | Starts the LibreOffice (port 2002) that the oracle and fixture generators talk to |
| `fixture_determinism.py` | Makes a generated workbook byte-identical on every run; `--check` lists the tracked workbooks that still depend on the moment they were made ([CONTRIBUTING.md](../CONTRIBUTING.md#generators-of-fixtures-are-deterministic)) |
| `package-wasm.sh` | Builds the npm package of `sheetrs-wasm` (`--demo` also puts it in `docs/pkg/`) |
| `build-wasm-rt.sh` | Builds the two base modules of the WebAssembly target of `sheetc`, next to `sheetc` |

## Measuring a change

```bash
cargo build --profile dist -p sheetrs-cli -p sheetrs-compiler
python3 scripts/bench.py fixtures DIR        # once; DIR is outside git
python3 scripts/bench.py run DIR lint,edit   # the groups the change touches
```

Run nothing else while it measures: a shared machine moves 10-20 % between sessions, and a number is compared with
the row of the same run, not with another session. When a metric moves beyond the spread of its runs, the change
updates the table of its flow in [docs/performance.md](../docs/performance.md), and
[PROFILING.md](../sheetrs-linter/docs/PROFILING.md) when it targets the cost of a rule or a phase. To see where the
time of a tool goes, sample it (a virtual machine has no hardware counters):

```bash
perf record -F 499 -e cpu-clock --call-graph fp ./target/fp/profiling/sheetrs lint -c scripts/benchmark.toml -i <file> > /dev/null
perf report --children --sort symbol -g none
```

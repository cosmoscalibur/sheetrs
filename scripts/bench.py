#!/usr/bin/env python3
"""The benchmarks of SheetRS, reproducible from the repository: the workbooks are generated, none is private.

    scripts/bench.py fixtures DIR [--rows 300000]
    scripts/bench.py run DIR [GROUP[,GROUP]...] [--runs 3] [--bin target/dist] [--python VENV_PYTHON] [--scenarios NAME[,NAME]...]

`fixtures` writes into DIR (outside git; every file can be regenerated):

  formulas.xlsx  sheetrs-compiler/scripts/gen_large_fixture.py with ROWS rows of `Calc`: 8 formula columns
                 (arithmetic, IF, exact VLOOKUP, running total, text assembly), `Summary` with five aggregates,
                 `Data`, `Lookup`, `settings` and one defined name; 300000 rows are 2.4 M formulas, 3.6 M cells
  formulas.ods   the same workbook converted by LibreOffice headless (`soffice --convert-to ods`)
  data.xlsx      sheetrs-compiler/scripts/gen_data_fixture.py: 600000 rows of numbers, labels and a row SUM
  data_declared.xlsx   the same with a `settings` sheet that declares inputs and outputs
  small.xlsx     sheetrs-compiler/tests/assets/loan.xlsx (a loan schedule: tens of formulas)

`run` measures each group with dist-profile binaries (`cargo build --profile dist -p sheetrs-cli -p
sheetrs-compiler`). Every command runs RUNS times; the table gives the minimum and the maximum of the wall
time and of the peak resident memory (`ru_maxrss` of the child: no `/usr/bin/time` is needed). Groups:

  lint stats query edit calculate   `sheetrs lint|stats|query|edit` and `sheetc calculate` on formulas.xlsx / .ods
  compiler                          `sheetc build` and the executable it writes, by stage (measure_perf.py)
  wasm                              the WebAssembly package (needs scripts/package-wasm.sh and scripts/build-wasm-rt.sh): time and
                                    peak memory of one call per tool, sizes, and
                                    the module of a compiled workbook against its native executable
  model                             the model use (inputs -> ranges): one call and 100 calls, SheetRS against LibreOffice
  python                            the Python `Book` (needs a wheel built with `maturin develop --profile dist`)
  libreoffice                       the same flows in LibreOffice through UNO (measure_libreoffice_flow.py)

The tables are printed as Markdown. Run nothing else on the machine while it measures: the numbers of a
shared machine move 10-20 % between sessions, so compare the rows of one run.
"""
import argparse
import hashlib
import os
import shutil
import statistics
import subprocess
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _workspace_version():
    """The version of the workspace (`Cargo.toml`), which names the npm tarball."""
    with open(os.path.join(ROOT, "Cargo.toml")) as f:
        for line in f:
            if line.startswith("version = "):
                return line.split('"')[1]


VERSION = _workspace_version()
SOFFICE = shutil.which("soffice") or "/opt/libreoffice26.8/program/soffice"


def run_once(cmd, stdin=None, env=None, cwd=None):
    """One execution: wall seconds, peak resident MB, exit code, standard output digest."""
    t0 = time.perf_counter()
    p = subprocess.Popen(cmd, stdin=subprocess.PIPE if stdin is not None else subprocess.DEVNULL,
                         stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=dict(os.environ, **(env or {})), cwd=cwd)
    if stdin is not None:
        p.stdin.write(stdin.encode())
        p.stdin.close()
    digest = hashlib.md5()
    head = b""
    while chunk := p.stdout.read(1 << 20):
        digest.update(chunk)
        if len(head) < (1 << 16):
            head += chunk[: (1 << 16) - len(head)]
    err = p.stderr.read().decode(errors="replace")
    _, status, ru = os.wait4(p.pid, 0)
    return {"wall": time.perf_counter() - t0, "rss": ru.ru_maxrss / 1024, "code": os.waitstatus_to_exitcode(status),
            "md5": digest.hexdigest(), "out": head.decode(errors="replace"), "err": err}


def measure(cmd, runs, ok=(0,), stdin=None, env=None, before=None, cwd=None):
    rs = []
    for _ in range(runs):
        if before:
            before()
        r = run_once(cmd, stdin, env, cwd)
        if r["code"] not in ok:
            sys.exit(f"failed (exit {r['code']}): {' '.join(map(str, cmd))}\n{r['err'][-400:]}")
        rs.append(r)
    return rs


def span(values, fmt):
    lo, hi = min(values), max(values)
    return fmt % lo if fmt % lo == fmt % hi else f"{fmt % lo}-{fmt % hi}"


def row(label, rs, extra=""):
    walls = [r["wall"] for r in rs]
    print(f"| {label} | {span(walls, '%.3f' if max(walls) < 0.1 else '%.2f')} s | {span([r['rss'] for r in rs], '%.0f')} MB | {extra} |",
          flush=True)


def table(title, header="Flow"):
    print(f"\n### {title}\n\n| {header} | Wall time | Peak memory | Output (md5 of the standard output / note) |\n|---|---|---|---|",
          flush=True)


def fixtures(args):
    d = os.path.abspath(args.dir)
    os.makedirs(d, exist_ok=True)
    scripts = os.path.join(ROOT, "sheetrs-compiler", "scripts")
    xlsx = os.path.join(d, "formulas.xlsx")
    subprocess.run([sys.executable, os.path.join(scripts, "gen_large_fixture.py"), xlsx, str(args.rows)], check=True)
    subprocess.run([sys.executable, os.path.join(scripts, "gen_data_fixture.py"), os.path.join(d, "data.xlsx")], check=True)
    subprocess.run([sys.executable, os.path.join(scripts, "gen_data_fixture.py"), os.path.join(d, "data_declared.xlsx"),
                    "--declared"], check=True)
    shutil.copy(os.path.join(ROOT, "sheetrs-compiler", "tests", "assets", "loan.xlsx"), os.path.join(d, "small.xlsx"))
    with tempfile.TemporaryDirectory() as profile:
        subprocess.run([SOFFICE, f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "ods",
                        "--outdir", d, xlsx], check=True)
    # The same workbook in XLSX with the results LibreOffice computed: `formulas.xlsx` stores 0 as
    # every cached result (so a plan on it always has stale results), this one stores the true ones.
    with tempfile.TemporaryDirectory() as profile, tempfile.TemporaryDirectory() as out:
        subprocess.run([SOFFICE, f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "xlsx",
                        "--outdir", out, os.path.join(d, "formulas.ods")], check=True)
        shutil.move(os.path.join(out, "formulas.xlsx"), os.path.join(d, "formulas_cached.xlsx"))
    for f in sorted(os.listdir(d)):
        print(f"{f}\t{os.path.getsize(os.path.join(d, f)) / 1e6:.1f} MB")


def group_lint(a):
    table("`sheetrs lint` (`scripts/benchmark.toml`)", "Workbook")
    for f in ("formulas.xlsx", "formulas.ods"):
        rs = measure([a.sheetrs, "lint", "-c", os.path.join(ROOT, "scripts", "benchmark.toml"), "-i", os.path.join(a.dir, f)],
                     a.runs, ok=(0, 1))
        row(f, rs, rs[0]["md5"])


def group_stats(a):
    table("`sheetrs stats`", "Workbook")
    for f in ("formulas.xlsx", "formulas.ods"):
        rs = measure([a.sheetrs, "stats", "-i", os.path.join(a.dir, f), "--format", "json"], a.runs)
        row(f, rs, rs[0]["md5"])


def group_query(a):
    table("`sheetrs query`", "Workbook and query")
    cases = [("--list-sheets", ["--list-sheets"]), ("`--range Calc!A1:H1000`", ["--range", "Calc!A1:H1000"]),
             ("`--sheet Summary`", ["--sheet", "Summary"]),
             ("`--find K00` (values, every sheet)", ["--find", "K00"]),
             ("`--find` in formulas (`VLOOKUP`)", ["--find", "VLOOKUP", "--in", "formula"])]
    for f in ("formulas.xlsx", "formulas.ods"):
        for label, args in cases:
            rs = measure([a.sheetrs, "query", "-i", os.path.join(a.dir, f), *args], a.runs)
            row(f"{f} {label}", rs, rs[0]["md5"])


EDIT_SCENARIOS = [
    ("delete-sheets", ["--delete-sheets", "Lookup"]),
    ("delete-names", ["--delete-names", "base_rate"]),
    ("formula-to-value-sheet", ["--formula-to-value", "Calc"]),
    ("formula-to-value-columns-rows", ["--formula-to-value", "Calc!B:D", "Calc!1:1000"]),
    ("plan", ["--delete-sheets", "Lookup", "--delete-names", "base_rate", "--formula-to-value", "Calc"]),
    ("delete-rows", ["--delete-rows", "Data!10:20"]),
    ("delete-rows-convert", ["--delete-rows", "Data!10:20", "--formula-to-value", "Calc"]),
    ("delete-rows-convert-summary", ["--delete-rows", "Calc!3:10", "--formula-to-value", "Summary"]),
    ("delete-columns", ["--delete-columns", "Data!B:D"]),
    ("delete-rows-columns", ["--delete-rows", "Data!10:20", "--delete-columns", "Data!B:D"]),
    ("structure-plan", ["--delete-sheets", "Lookup", "--delete-names", "base_rate", "--delete-rows", "Data!10:20",
                        "--delete-columns", "Data!B:D"]),
]


def group_edit(a):
    table("`sheetrs edit` (a deletion ends with the engine calculating every formula)", "Workbook and scenario")
    wanted = set(a.scenarios.split(",")) if a.scenarios else None
    with tempfile.TemporaryDirectory() as work:
        for f in a.workbooks.split(","):
            ext = f.rsplit(".", 1)[1]
            out = os.path.join(work, f"out.{ext}")
            for name, args in EDIT_SCENARIOS:
                if wanted and name not in wanted:
                    continue
                rs = measure([a.sheetrs, "edit", "-i", os.path.join(a.dir, f), *args, "-o", out], a.runs,
                             before=lambda: os.path.exists(out) and os.remove(out))
                row(f"{f} {name}", rs, f"`{' '.join(args)}`")


def group_calculate(a):
    table("`sheetc calculate`", "Flow")
    f = os.path.join(a.dir, "formulas.xlsx")
    with tempfile.TemporaryDirectory() as work:
        out = os.path.join(work, "out.xlsx")
        rs = measure([a.sheetc, "calculate", "-i", f, "-o", out], a.runs, before=lambda: os.path.exists(out) and os.remove(out))
        row("every formula calculated, workbook written (`-o OUT`)", rs)
    rs = measure([a.sheetc, "calculate", "-i", f, "--range", "Summary!A1"], a.runs)
    row("one declared output as JSON (`--range Summary!A1`)", rs, rs[0]["md5"])


def group_compiler(a):
    perf = os.path.join(a.dir, "perf")
    os.makedirs(perf, exist_ok=True)
    for src, dst in (("small.xlsx", "small.xlsx"), ("data.xlsx", "large.xlsx"), ("data_declared.xlsx", "large_declared.xlsx"),
                     ("formulas.xlsx", "formulas.xlsx")):
        shutil.copy(os.path.join(a.dir, src), os.path.join(perf, dst))
    print("\n### `sheetc build` and the executable, by stage\n", flush=True)
    r = subprocess.run([sys.executable, os.path.join(ROOT, "sheetrs-compiler", "scripts", "measure_perf.py")],
                       env=dict(os.environ, PERF_DIR=perf, SHEETC=a.sheetc, REPS=str(a.runs)), check=True, text=True,
                       capture_output=True)
    print(r.stdout, flush=True)


PY_STEP = r'''
import sys, time
import sheetrs
path, out, step = sys.argv[1], sys.argv[2], sys.argv[3]
t = time.perf_counter(); book = sheetrs.Book(path); t_open = time.perf_counter() - t
t = time.perf_counter()
if step == "range": book.range(["Calc!A1:H1000"])
elif step == "find": book.find("K00")
elif step == "lint": book.lint()
elif step == "delete_rows": book.delete_rows("Calc!3:10").save(out)
elif step == "calculate": book.calculate({"rate": 0.07}, ["Summary!A1", "Summary!A3", "Summary!A4"]).value_ranges()
print(f"{t_open:.4f} {time.perf_counter() - t:.4f}")
'''


def group_python(a):
    """Every step in its own process, so that each has its own peak memory (the process opens the Book first)."""
    print("\n### Python `Book` on `formulas.xlsx` (one process per step and run)\n\n"
          "| Step | Wall time of the step | Wall time of the process | Peak memory of the process |\n|---|---|---|---|", flush=True)
    f = os.path.join(a.dir, "formulas.xlsx")
    steps = [("open", "`Book(path)` (reads the workbook once, for the statistics)"),
             ("range", "`range([\"Calc!A1:H1000\"])`"), ("find", "`find(\"K00\")`"), ("lint", "`lint()`"),
             ("delete_rows", "`delete_rows(\"Calc!3:10\").save(out)` (with the recalculation)"),
             ("calculate", "`calculate({\"rate\": 0.07}, [\"Summary!A1\", \"Summary!A3\", \"Summary!A4\"]).value_ranges()`")]
    with tempfile.TemporaryDirectory() as work:
        for step, label in steps:
            rs = measure([a.python, "-c", PY_STEP, f, os.path.join(work, "out.xlsx"), step], a.runs)
            times = [[float(x) for x in r["out"].split()] for r in rs]
            walls = [t[0] if step == "open" else t[1] for t in times]
            print(f"| {label} | {span(walls, '%.2f')} s | {span([r['wall'] for r in rs], '%.2f')} s | {span([r['rss'] for r in rs], '%.0f')} MB |",
                  flush=True)


def group_libreoffice(a):
    print("\n### The same flows in LibreOffice 26.8 through UNO\n", flush=True)
    r = subprocess.run(["/opt/libreoffice26.8/program/python", os.path.join(ROOT, "scripts", "measure_libreoffice_flow.py"),
                        os.path.join(a.dir, "formulas.xlsx"), "--rounds", str(a.runs), "--sheetrs", a.sheetrs, "--sheetc", a.sheetc],
                       check=True, text=True, capture_output=True)
    print(r.stdout, flush=True)


def group_wasm(a):
    """The WebAssembly package (sizes, the time of a call against the command) and the module of a compiled workbook."""
    npm = os.path.join(ROOT, "target", "npm")
    print("\n### The WebAssembly package (`scripts/package-wasm.sh`)\n\n| File | Bytes | gzip |\n|---|---:|---:|")
    for f in ("sheetrs_wasm_bg.wasm", f"sheetrs-wasm-{VERSION}.tgz"):
        path = os.path.join(npm, f)
        gz = len(subprocess.run(["gzip", "-c", path], capture_output=True, check=True).stdout)
        print(f"| {f} | {os.path.getsize(path):,} | {gz:,} |")
    for f in ("sheetrs-rt.wasm", "sheetrs-rt-template.wasm", "sheetc"):
        path = os.path.join(a.bin, f)
        gz = len(subprocess.run(["gzip", "-c", path], capture_output=True, check=True).stdout)
        print(f"| {f} (dist) | {os.path.getsize(path):,} | {gz:,} |")
    print("\nOne call of each tool (`sheetrs-wasm/tests/timing.mjs`, median of 20 after 3 warm-ups) against the command:\n", flush=True)
    r = subprocess.run(["node", os.path.join(ROOT, "sheetrs-wasm", "tests", "timing.mjs")], check=True, text=True, capture_output=True,
                       env=dict(os.environ, BIN=a.bin))
    print(r.stdout, flush=True)
    table("Peak memory of one call (one process per scenario: Node loads the package, makes the call once and exits; `none` stops after loading)",
          "Call")
    calls = {"lint": ("lint", ["-i", "tests/lint_book.xlsx", "--format", "sarif"]),
             "stats": ("stats", ["-i", "tests/lint_book.xlsx", "--format", "json"]),
             "query": ("query", ["-i", "sheetrs-edit/tests/assets/book/book.xlsx", "--range", "Data!A1:C9"]),
             "edit": ("edit", ["-i", "sheetrs-edit/tests/assets/book/book.xlsx", "--delete-rows", "Data!10:11", "--formula-to-value",
                               "Data!A:A", "--delete-columns", "Data!B", "-o", "{out}"]),
             "calculate": ("sheetc", ["calculate", "-i", "sheetrs-compiler/tests/assets/loan.xlsx", "--request", "-"])}
    request = '{"inputs":{"principal":250000,"rate":0.045,"years":15,"month":24}}'
    caller = os.path.join(ROOT, "sheetrs-wasm", "tests", "call_run.mjs")
    base = measure(["node", caller, "none"], a.runs)
    row("Node loading the package, no call (`none`)", base)
    with tempfile.TemporaryDirectory() as work:
        for name, (tool, args) in calls.items():
            args = [x.replace("{out}", os.path.join(work, "out.xlsx")) if "{out}" in x else x for x in args]
            args = [os.path.join(ROOT, x) if x.startswith(("tests/", "sheetrs-")) else x for x in args]
            wasm_rs = measure(["node", caller, name], a.runs, cwd=ROOT)
            exe = [a.sheetc if tool == "sheetc" else a.sheetrs] + ([] if tool == "sheetc" else [tool]) + args
            cmd_rs = measure(exe, a.runs, ok=(0, 1), stdin=request if name == "calculate" else None)
            row(f"{name}: WebAssembly call (process of Node)", wasm_rs)
            row(f"{name}: command", cmd_rs)
    table("A compiled workbook: the native executable against its WebAssembly module (whole process, Node 22)", "Program and request")
    with tempfile.TemporaryDirectory() as work:
        for name, f, req in (("small.xlsx", "small.xlsx", '{"ranges":["Loan!B4"]}'),
                             ("formulas.xlsx", "formulas.xlsx", '{"ranges":["Summary!A1","Summary!A2","Summary!A3"]}')):
            exe, wasm = os.path.join(work, name + ".exe"), os.path.join(work, name + ".wasm")
            subprocess.run([a.sheetc, "build", "-i", os.path.join(a.dir, f), "-o", exe], check=True, capture_output=True)
            subprocess.run([a.sheetc, "build", "-i", os.path.join(a.dir, f), "--target", "wasm", "-o", wasm], check=True,
                           capture_output=True)
            nat = measure([exe, "--request", "-"], a.runs, stdin=req)
            mod = measure(["node", os.path.join(ROOT, "sheetrs-wasm", "tests", "module_run.mjs"), wasm, req], a.runs)
            row(f"{name}: native executable ({os.path.getsize(exe):,} B)", nat, nat[0]["md5"])
            row(f"{name}: WebAssembly module ({os.path.getsize(wasm):,} B)", mod, mod[0]["md5"])
            os.remove(exe)
            os.remove(wasm)


MODELS = [
    # workbook, input name and cell, base value, output names and cells
    ("data_declared.xlsx", "scale", "Summary!B4", 2.0, ["total", "scaled"], ["Summary!B1", "Summary!B5"]),
    ("formulas.xlsx", "rate", "Data!I1", 0.05, ["total", "month3", "keys"], ["Summary!A1", "Summary!A3", "Summary!A4"]),
]
CALLS = 100


def group_model(a):
    """The model use: inputs -> ranges. One call, and 100 calls with the model kept (served, or document open)."""
    import json
    for f, name, cell, base, ranges, cells in MODELS:
        path = os.path.join(a.dir, f)
        print(f"\n### The model use on `{f}`: input `{name}` ({cell}), ranges {', '.join('`%s`' % r for r in ranges)}\n\n"
              f"| Flow | One call, wall time | Peak memory | {CALLS} calls |\n|---|---|---|---|", flush=True)
        request = json.dumps({"inputs": {name: base}, "ranges": ranges})
        rs = measure([a.sheetc, "calculate", "-i", path, "--request", request], a.runs)
        print(f"| `sheetc calculate --request` (compiles in memory, every call) | {span([r['wall'] for r in rs], '%.2f')} s | "
              f"{span([r['rss'] for r in rs], '%.0f')} MB | {CALLS} separate runs: no state is kept |", flush=True)
        with tempfile.TemporaryDirectory() as work:
            exe = os.path.join(work, "model.exe")
            rs = measure([a.sheetc, "build", "-i", path, "-o", exe], a.runs)
            print(f"| `sheetc build` of the executable (once; LibreOffice has no equivalent) | {span([r['wall'] for r in rs], '%.2f')} s | "
                  f"{span([r['rss'] for r in rs], '%.0f')} MB | |", flush=True)
            rs = measure([exe, "--request", "-"], a.runs, stdin=request)
            one = min(r["wall"] for r in rs)
            print(f"| compiled executable, `--request` | {span([r['wall'] for r in rs], '%.3f')} s | "
                  f"{span([r['rss'] for r in rs], '%.0f')} MB | {CALLS} separate runs |", flush=True)
            lines = "\n".join(json.dumps({"inputs": {name: base * (1 + 0.001 * i)}, "ranges": ranges}) for i in range(CALLS)) + "\n"
            rs = measure([exe, "--serve"], a.runs, stdin=lines)
            print(f"| compiled executable, `--serve` | (one call, as above) | {span([r['rss'] for r in rs], '%.0f')} MB | "
                  f"{span([r['wall'] for r in rs], '%.3f')} s in all, {min(r['wall'] for r in rs) * 10:.3f} ms per call, the start of the process included |", flush=True)
        code = ("import sys, json, sheetrs\npath = sys.argv[1]\nsheetrs.Book(path).calculate(%s, %s).value_ranges()\n"
                % (json.dumps({name: base}), json.dumps(ranges)))
        rs = measure([a.python, "-c", code, path], a.runs)
        print(f"| Python `Book(path).calculate(inputs, ranges).value_ranges()` | {span([r['wall'] for r in rs], '%.2f')} s | "
              f"{span([r['rss'] for r in rs], '%.0f')} MB | {CALLS} separate runs: each call opens the workbook again |", flush=True)
        lo = subprocess.run(["/opt/libreoffice26.8/program/python", os.path.join(ROOT, "scripts", "measure_libreoffice_flow.py"), path,
                             "--rounds", str(a.runs), "--only", "libreoffice-model", "--model-input", cell, "--model-ranges",
                             ",".join(cells), "--model-base", str(base), "--calls", str(CALLS)],
                            check=True, text=True, capture_output=True).stdout
        rounds = [json.loads(l.split(" ", 1)[1]) for l in lo.splitlines() if l.startswith("MODEL_JSON")]
        first = [r["calls"][0] for r in rounds]
        loads = [r["load"] for r in rounds]
        rest = [sum(r["calls"][1:]) / (len(r["calls"]) - 1) for r in rounds]
        total = [sum(r["calls"]) for r in rounds]
        print(f"| LibreOffice through UNO: open once, set the input, `calculateAll()`, read the cells | load {span(loads, '%.2f')} s + "
              f"{span(first, '%.2f')} s | {span([r['hwm_mb'] for r in rounds], '%.0f')} MB | {span(total, '%.1f')} s for the {CALLS} calls "
              f"(the first with `calculateAll()`, the next with `calculate()`: {span([x * 1000 for x in rest], '%.1f')} ms per call) |", flush=True)


GROUPS = {"lint": group_lint, "stats": group_stats, "query": group_query, "edit": group_edit, "calculate": group_calculate,
          "compiler": group_compiler, "wasm": group_wasm, "model": group_model, "python": group_python, "libreoffice": group_libreoffice}

if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    f = sub.add_parser("fixtures")
    f.add_argument("dir")
    f.add_argument("--rows", type=int, default=300_000)
    r = sub.add_parser("run")
    r.add_argument("dir")
    r.add_argument("groups", nargs="?", default=",".join(GROUPS))
    r.add_argument("--runs", type=int, default=3)
    r.add_argument("--bin", default=os.path.join(ROOT, "target", "dist"))
    r.add_argument("--python", default=sys.executable)
    r.add_argument("--workbooks", default="formulas.xlsx,formulas.ods",
                   help="the `edit` group: these workbooks of DIR (comma separated)")
    r.add_argument("--scenarios", help="the `edit` group: only these scenarios, by name (comma separated)")
    a = ap.parse_args()
    if a.cmd == "fixtures":
        fixtures(a)
    else:
        a.dir = os.path.abspath(a.dir)
        a.sheetrs, a.sheetc = os.path.join(a.bin, "sheetrs"), os.path.join(a.bin, "sheetc")
        for g in a.groups.split(","):
            GROUPS[g](a)

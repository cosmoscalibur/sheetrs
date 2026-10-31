"""Performance matrix of sheetc: for each workbook size, the compiler build (total and by stage),
the engine start-up (`--meta`) and the engine execution of one request; wall time, user and
system CPU, peak memory and executable size. Median of REPS runs.

PERF_DIR holds small.xlsx, delgado.xlsx, large.xlsx (49 MB, gen_data_fixture.py), large_declared.xlsx
(the same data plus a settings sheet: the build keeps only what the outputs need) and
formulas.xlsx (14 MB, 1.2 M formulas, gen_large_fixture.py); the same files in every measurement
(tests/production/README.local.txt). delgado.toml is the config of the Delgado workbook.
SHEETC points to the binary (default target/release/sheetc). Output: matrix on stdout and
PERF_DIR/matrix.json. Flamegraph: build with `--features profile`, SHEETC_FLAMEGRAPH=out.svg."""
import os, re, subprocess, sys, time, json, statistics

SC = os.environ.get("SHEETC", "target/release/sheetc")
REPS = int(os.environ.get("REPS", "3"))


def run(cmd, stdin=None, env=None):
    t = time.perf_counter()
    e = dict(os.environ, **(env or {}))
    p = subprocess.Popen(cmd, stdin=subprocess.PIPE if stdin is not None else None,
                         stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=e)
    if stdin is not None:
        p.stdin.write(stdin.encode()); p.stdin.close()
    out = p.stdout.read(); err = p.stderr.read()
    _, status, ru = os.wait4(p.pid, 0)
    return dict(wall=time.perf_counter() - t, user=ru.ru_utime, sys=ru.ru_stime,
                rss=ru.ru_maxrss / 1024, code=os.waitstatus_to_exitcode(status),
                out=out, err=err.decode(errors="replace"))


def med(rs, k):
    return statistics.median(r[k] for r in rs)


def bench(cmd, stdin=None, reps=REPS, env=None):
    rs = [run(cmd, stdin, env) for _ in range(reps)]
    r = {k: med(rs, k) for k in ("wall", "user", "sys", "rss")}
    r["code"] = rs[0]["code"]; r["err"] = rs[0]["err"]
    return r


def checked(r, what):
    """A measurement of a failed command is not a measurement: stop with its message."""
    if r["code"] != 0:
        sys.exit(f"{what} failed (exit {r['code']}): {r['err'].strip()[:300]}")
    return r


def seconds(txt):
    m = re.match(r"\s*([\d.]+)(ms|µs|s|ns)", txt)
    v = float(m.group(1))
    return v * {"s": 1, "ms": 1e-3, "µs": 1e-6, "ns": 1e-9}[m.group(2)]


def stages(err):
    """Stage seconds from a SHEETC_TRACE build (compile marks are cumulative)."""
    st = {}
    marks = {}
    for line in err.splitlines():
        m = re.match(r"\[(stage|compile)\s+(\S+)\]\s+(.*)", line)
        if not m:
            continue
        t = seconds(m.group(2))
        if m.group(1) == "stage":
            st[m.group(3)] = t
        else:
            marks[m.group(3)] = t
    seq = ["parsed formulas", "bindings", "index built", "cone + csr", "scc", "taint", "fold evaluated", "needed"]
    prev = 0.0
    d = {}
    for k in seq:
        if k in marks:
            d[k] = marks[k] - prev
            prev = marks[k]
    st["lower formulas"] = d.get("parsed formulas", 0)
    st["dependency graph"] = sum(d.get(k, 0) for k in ("bindings", "index built", "cone + csr", "scc", "taint"))
    st["fold (engine at build)"] = d.get("fold evaluated", 0)
    st["assemble"] = d.get("needed", 0)
    return st


FILES = [
    ("small", "small.xlsx", '{"ranges":["Loan!B1:B8"]}', None),
    ("delgado", "delgado.xlsx", '{"ranges":["OUTPUT!D457:E462"]}', "delgado.toml"),
    ("large", "large.xlsx", '{"ranges":["Summary!B1:B5"]}', None),
    ("large (declared)", "large_declared.xlsx", '{"inputs":{"scale":2}}', "declared"),
    ("formulas", "formulas.xlsx", '{"ranges":["Summary!A1:A5"]}', None),
]

if __name__ == "__main__":
    d = os.environ.get("PERF_DIR", ".")
    res = {}
    for name, f, req, cfg in FILES:
        exe = f"{d}/{name.split()[0]}_{cfg or 'any'}.exe"
        declared = cfg == "declared"
        extra = ["--config", f"{d}/{cfg}"] if cfg and not declared and os.path.exists(f"{d}/{cfg}") else []
        cmd = [SC, "build", f"{d}/{f}", "-o", exe, "--outputs", "declared" if declared else "any"] + extra
        r = {"input_mb": os.path.getsize(f"{d}/{f}") / 1e6}
        r["build"] = checked(bench(cmd), f"{name}: build")
        r["exe_mb"] = os.path.getsize(exe) / 1e6
        r["stages"] = stages(run(cmd, env={"SHEETC_TRACE": "1"})["err"])
        r["meta"] = checked(bench([exe, "--meta"]), f"{name}: --meta")
        r["exec"] = checked(bench([exe], stdin=req), f"{name}: request {req}")
        res[name] = r
    json.dump(res, open(d + "/matrix.json", "w"), indent=1)
    cols = ["wall", "user", "sys", "rss"]
    print("| file | stage | wall s | user s | sys s | peak MB | size MB |")
    print("|---|---|--:|--:|--:|--:|--:|")
    for name, r in res.items():
        b = r["build"]
        print(f'| {name} | input workbook | | | | | {r["input_mb"]:.2f} |')
        print(f'| {name} | compiler: build (total) | {b["wall"]:.2f} | {b["user"]:.2f} | {b["sys"]:.2f} | {b["rss"]:.0f} | {r["exe_mb"]:.2f} |')
        for k in ("read workbook", "lower formulas", "dependency graph", "fold (engine at build)", "assemble", "write executable"):
            if k in r["stages"]:
                print(f'| {name} | &nbsp;&nbsp;{k} | {r["stages"][k]:.2f} | | | | |')
        rest = b["wall"] - sum(r["stages"].get(k, 0) for k in ("read workbook", "lower formulas", "dependency graph", "fold (engine at build)", "assemble", "write executable"))
        print(f'| {name} | &nbsp;&nbsp;other (compile remainder, exit) | {max(rest, 0):.2f} | | | | |')
        for label, key in (("engine: start-up (--meta)", "meta"), ("engine: execute request", "exec")):
            x = r[key]
            print(f'| {name} | {label} | {x["wall"]:.3f} | {x["user"]:.3f} | {x["sys"]:.3f} | {x["rss"]:.0f} | |')

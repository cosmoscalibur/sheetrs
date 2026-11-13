#!/usr/bin/env python3
"""Which arguments of a function take an array written in the formula one element at a time (`ABS({-1;-2})`
is an array of two numbers) and which take the array whole (`SUM({1;2})` is one number), in
`LibreOffice`, and the table the engine takes from it (`sheetrs-engine/src/functions/lift_table.rs`).

    scripts/probe_lift.py ask RAW.json        # asks LibreOffice 26.8 (needs soffice, openpyxl, cargo)
    scripts/probe_lift.py verify RAW.json     # asks for each candidate the element of the result against the call per element
    scripts/probe_lift.py table RAW.json      # writes lift_table.rs from the answers

An array in a place where a function takes one value, in a formula that is not an array formula, is
not the first of its elements: the function is called for each element and gives an array (`ABS`,
`ROUND`, `IF`, `CHOOSE`, `INDEX` with an array of rows, `VLOOKUP` with an array of keys…). A function that
takes an array there (`SUM`, `MAX`, `AND`, `INDEX`'s own array, `VLOOKUP`'s table) gets it whole. The
test is the size of the result: `ROWS(F(…{2;3}…))` is 2 for the first and `Err:504` (a number is not
an array) for the second. Each function is called with the array in each position, over numbers, over
a text, and with ranges in the other positions that take them.
"""
import json
import os
import sys
from concurrent.futures import ThreadPoolExecutor

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import lo_probe  # noqa: E402
import probe_arity as pa  # noqa: E402

ROOT = os.path.join(HERE, "..", "..")
OUT = os.path.join(ROOT, "sheetrs-engine", "src", "functions", "lift_table.rs")
NUM = "{2;3}"
TXT = '{"ab";"cd"}'
BASES = [
    lambda n: ["2"] * n,
    lambda n: ["Data!A1:A8"] + ["2"] * (n - 1),
    lambda n: ['"abc"'] + ["2"] * (n - 1),
    lambda n: ["2", "Data!A1:A8"] + ["2"] * (n - 2) if n >= 2 else ["2"] * n,
    lambda n: ["3", "2", "1"] + ["1"] * (n - 3) if n >= 3 else ["3"] * n,
    lambda n: ["2", "Data!A1:C8", "2"] + ["0"] * (n - 3) if n >= 3 else ["2"] * n,
]


def rows(funcs, spelling):
    out = []
    for ns, name, mn, mx in funcs:
        if (ns, name) not in spelling:
            continue
        prefix, nm = spelling[(ns, name)]
        top = min(mx, max(mn, 1) + 2, 6)
        for n in range(max(mn, 1), top + 1):
            for base in BASES:
                args = base(n)[:n]
                for p in range(n):
                    for arr in (NUM, TXT):
                        a = list(args)
                        a[p] = arr
                        out.append((ns, name, n, p, f"ROWS({prefix}{nm}({','.join(a)}))"))
    return out


def ask(out_path):
    funcs = [(ns, name, int(mn), int(mx)) for ns, name, mn, mx in pa.registry()]
    funcs = [f for f in funcs if (f[0], f[1]) not in pa.NOT_PROBED and not f[1].startswith("_")]
    spell_path = out_path + ".spelling"
    if os.path.exists(spell_path):
        spelling = {tuple(k.split("|")): tuple(v) for k, v in json.load(open(spell_path)).items()}
    else:
        spelling = pa.find_spellings(funcs)
        json.dump({"|".join(k): list(v) for k, v in spelling.items()}, open(spell_path, "w"))
    rs = rows(funcs, spelling)
    seen, uniq = set(), []
    for r in rs:
        if r[4] not in seen:
            seen.add(r[4])
            uniq.append(r)
    rs = uniq
    allf = [r[4] for r in rs]
    print(len(allf), "formulas", file=sys.stderr)
    chunks = [allf[i : i + 800] for i in range(0, len(allf), 800)]

    def run(chunk):
        try:
            return lo_probe.probe(chunk)
        except Exception:
            res = []
            for f in chunk:
                try:
                    res.append(lo_probe.probe([f])[0])
                except Exception:
                    res.append(("crash", ""))
            return res

    with ThreadPoolExecutor(3) as pool:
        answers = list(pool.map(run, chunks))
    flat = [a for ans in answers for a in ans]
    json.dump({"rows": rs, "answers": flat}, open(out_path, "w"))


def candidates(raw):
    """Up to four samples for each argument of each function whose result has 2 rows (the call over
    the array of two elements): the result is the array of the calls only if its elements are what the
    calls over each element give (`INDEX({2;3})` has two rows and is not one)."""
    per = {}
    for (ns, name, n, p, f), (kind, val) in zip(raw["rows"], raw["answers"]):
        if kind == "float" and val == "2":
            got = per.setdefault((name, ns, p), [])
            if len(got) < 4:
                got.append(f)
    out = []
    for (name, ns, p), fs in per.items():
        for f in fs:
            call = f[len("ROWS("):-1]
            for arr, (e1, e2) in ((NUM, ("2", "3")), (TXT, ('"ab"', '"cd"'))):
                if arr in call:
                    out.append((name, ns, p, call, call.replace(arr, e1, 1), call.replace(arr, e2, 1), arr))
    return out


def verify(raw_path):
    raw = json.load(open(raw_path))
    cands = candidates(raw)
    fs = sorted({x for c in cands for x in (f"INDEX({c[3]},1,1)", f"INDEX({c[3]},2,1)", f"COLUMNS({c[3]})", c[4], c[5])})
    print(len(cands), "candidates,", len(fs), "formulas", file=sys.stderr)
    chunks = [fs[i : i + 800] for i in range(0, len(fs), 800)]

    def run(chunk):
        try:
            return lo_probe.probe(chunk)
        except Exception:
            res = []
            for f in chunk:
                try:
                    res.append(lo_probe.probe([f])[0])
                except Exception:
                    res.append(("crash", ""))
            return res

    with ThreadPoolExecutor(3) as pool:
        answers = list(pool.map(run, chunks))
    flat = {f: a for chunk, ans in zip(chunks, answers) for f, a in zip(chunk, ans)}
    json.dump(flat, open(raw_path + ".verify", "w"))


def table(raw_path):
    raw = json.load(open(raw_path))
    ver = json.load(open(raw_path + ".verify"))
    funcs = {(n, ns): int(mx) for ns, n, mn, mx in pa.registry()}
    lifts = {}
    bad = set()
    for name, ns, p, call, c1, c2, arr in candidates(raw):
        # The elements of the result are what the call over each element gives, and the result is one
        # column (`HSTACK({2;3},2)` is not).
        a1, a2 = ver.get(f"INDEX({call},1,1)"), ver.get(f"INDEX({call},2,1)")
        b1, b2 = ver.get(c1), ver.get(c2)
        cols = ver.get(f"COLUMNS({call})")
        if not all((a1, a2, b1, b2, cols)) or "crash" in (a1[0], b1[0]):
            continue
        # An element of an array is a number: a boolean the call gives is 1 or 0 once picked out.
        norm = lambda a: ("float", "1" if a[1] == "true" else "0") if a[0] == "boolean" else tuple(a)
        both_values = a1[0] != "error" and b1[0] != "error" and a2[0] != "error" and b2[0] != "error"
        # What a function gives that is a date or a time reads as a number once picked out of an array
        # (the format is lost), so only that it is a number is compared.
        shown = {"date", "time", "currency", "percentage"}
        def same_value(x, y):
            if y[0] in shown and x[0] == "float":
                return True
            return norm(x) == norm(y)
        if cols == ["float", "1"] and same_value(a1, b1) and same_value(a2, b2):
            lifts.setdefault((name, ns), set()).add(p)
        elif both_values:
            bad.add((name, ns, p))
    for name, ns, p in bad:
        if (name, ns) in lifts:
            lifts[(name, ns)].discard(p)
    lifts = {k: v for k, v in lifts.items() if v}
    for (name, ns), ps in list(lifts.items()):
        # A function with any number of arguments that takes the array of every position sampled does
        # it for every later one.
        if funcs.get((name, ns)) == 255 and len(ps) >= 3:
            lifts[(name, ns)] = ps | {i for i in range(max(ps) + 1, 32)}
    entries = sorted((name, ns, sum(1 << p for p in ps)) for (name, ns), ps in lifts.items())
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("//! The arguments of each function that take an array written in the formula one element at a\n")
        f.write("//! time, from `LibreOffice` 26.8.\n//!\n")
        f.write("//! Generated by `sheetrs-compiler/scripts/probe_lift.py`: do not edit.\n\n")
        f.write("/// `(name, namespace, arguments)`, sorted by name and namespace: a bit per argument, the first is bit 0.\n")
        f.write("pub(super) const LIFT: &[(&str, &str, u32)] = &[\n")
        for name, ns, mask in entries:
            f.write(f'    ("{name}", "{ns}", {"u32::MAX" if mask == 0xFFFFFFFF else mask}),\n')
        f.write("];\n")
    print(len(entries), "functions written", file=sys.stderr)


if __name__ == "__main__":
    {"ask": lambda: ask(sys.argv[2]), "verify": lambda: verify(sys.argv[2]), "table": lambda: table(sys.argv[2])}[sys.argv[1]]()

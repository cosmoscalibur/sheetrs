#!/usr/bin/env python3
"""What a function that takes one value does with an array that `IF` chose element by element (a
jump matrix), outside array evaluation, in `LibreOffice` 26.8, and the table the engine takes from it
the rule the engine follows (`Array::jump`, `Engine::call`).

    scripts/probe_jump.py ask RAW.json        # asks LibreOffice (needs soffice, openpyxl, cargo)
    scripts/probe_jump.py check RAW.json      # says what each function did

`ABS(IF({1;0},2,3))` in a cell is the call over the first element, 2, and `SUM(ABS(IF({1;1},2,3)))` is
2, not 4: a function does not take such an array element by element unless the evaluation is an array
evaluation (`SUMPRODUCT`, an array formula). Every function gets the first element. For every argument
that `lift_table.rs` says a function takes one element at a time, the call is made with `IF({1;0},2,3)`
there and with `2`: the answers are the same for all of them (an answer that differs is listed).
"""
import json
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import lo_probe  # noqa: E402
import probe_arity as pa  # noqa: E402

ROOT = os.path.join(HERE, "..", "..")
LIFT = os.path.join(ROOT, "sheetrs-engine", "src", "functions", "lift_table.rs")
JUMP = "IF({1;0},2,3)"
BASES = [
    lambda n: ["2"] * n,
    lambda n: ['"abc"'] + ["2"] * (n - 1),
    lambda n: ["3", "2", "1"] + ["1"] * (n - 3) if n >= 3 else ["3"] * n,
]


def lifts():
    out = {}
    for name, ns, mask in re.findall(r'\("([^"]+)", "([^"]+)", ([^)]+)\),', open(LIFT, encoding="utf-8").read()):
        out[(name, ns)] = 0xFFFFFFFF if mask == "u32::MAX" else int(mask)
    return out


def ask(out_path):
    table = lifts()
    funcs = [(ns, name, int(mn), int(mx)) for ns, name, mn, mx in pa.registry()]
    funcs = [f for f in funcs if (f[1], f[0]) in table and (f[0], f[1]) not in pa.NOT_PROBED and not f[1].startswith("_")]
    spelling = pa.find_spellings(funcs)
    rows = []
    for ns, name, mn, mx in funcs:
        if (ns, name) not in spelling:
            continue
        prefix, nm = spelling[(ns, name)]
        mask = table[(name, ns)]
        for p in range(min(max(mx, 1), 6)):
            if not mask >> p & 1:
                continue
            for n in range(max(mn, p + 1), min(mx, max(mn, p + 1) + 2) + 1):
                for base in BASES:
                    args = base(n)[:n]
                    j = list(args)
                    j[p] = JUMP
                    s = list(args)
                    s[p] = "2"
                    rows.append((ns, name, n, p, f"{prefix}{nm}({','.join(j)})", f"{prefix}{nm}({','.join(s)})"))
    seen, uniq = set(), []
    for r in rows:
        if r[4] not in seen:
            seen.add(r[4])
            uniq.append(r)
    rows = uniq
    allf = [f for r in rows for f in (r[4], r[5])]
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
    json.dump({"rows": rows, "answers": flat}, open(out_path, "w"))


def classify(raw):
    """`{(name, ns): {position: 'first' | 'refused' | 'other'}}`."""
    res = {}
    for i, (ns, name, n, p, jf, sf) in enumerate(raw["rows"]):
        j, s = raw["answers"][2 * i], raw["answers"][2 * i + 1]
        if "crash" in (j[0], s[0]) or s[0] == "error":
            continue
        norm = lambda a: ("float", "1" if a[1] == "true" else "0") if a[0] == "boolean" else tuple(a)
        if norm(j) == norm(s):
            what = "first"
        elif j[0] == "error" and j[1] == "Err:504":
            what = "refused"
        else:
            what = "other"
        res.setdefault((name, ns), {}).setdefault(p, []).append(what)
    out = {}
    for key, pos in res.items():
        for p, whats in pos.items():
            # Several samples of the same argument: refused if any sample was, `first` if any was and
            # none refused.
            out.setdefault(key, {})[p] = "refused" if "refused" in whats else ("first" if "first" in whats else "other")
    return out


def check(raw_path):
    raw = json.load(open(raw_path))
    cls = classify(raw)
    counts = {}
    for pos in cls.values():
        for w in pos.values():
            counts[w] = counts.get(w, 0) + 1
    other = sorted((name, ns, p) for (name, ns), pos in cls.items() for p, w in pos.items() if w != "first")
    print(counts, "not the first element:", other)


if __name__ == "__main__":
    {"ask": lambda: ask(sys.argv[2]), "check": lambda: check(sys.argv[2])}[sys.argv[1]]()

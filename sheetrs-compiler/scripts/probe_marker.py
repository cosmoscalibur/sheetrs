#!/usr/bin/env python3
"""How a function that takes one value reads the element that the missing `else` of an `IF` gave
(`IF({0;1},1)`: the element of the first row) in `LibreOffice` 26.8, and the table the engine takes from
it (`sheetrs-engine/src/functions/marker_table.rs`).

    scripts/probe_marker.py ask RAW.json        # asks LibreOffice (needs soffice, openpyxl, cargo)
    scripts/probe_marker.py table RAW.json      # writes marker_table.rs from the answers

The element is read as the number 0 where a number is taken (`ABS`, `+`), as the text `FALSE` where a text
is (`LEN` is 5, `&` gives `FALSE`), as an empty text in some functions (`T`), and an add-in refuses it
(`ISEVEN` of a text and of this element: `Err:504`, the same). For every argument that `lift_table.rs` says a function takes one element at a time,
the call is made with `IF({0;1},2)` there and with each of `0`, `"FALSE"`, `""` and `FALSE()`: the first of
these that gives the same answer is the way the function reads it (the number 0 when more than one does,
for it is the one that matters least). A call that is `Err:504` for the element and for no candidate is a
refusal of the element itself: none was found (the add-ins that give `Err:504` for the element do for the text too).
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
OUT = os.path.join(ROOT, "sheetrs-engine", "src", "functions", "marker_table.rs")
MARKER = "IF({0;1},2)"
# Their answers are random: the same error for every reading says nothing about the element.
RANDOM = {"RAND", "RANDBETWEEN", "RANDARRAY"}
CANDIDATES = [("number", "0"), ("text", '"FALSE"'), ("empty", '""'), ("bool", "FALSE()")]
BASES = [
    lambda n: ["2"] * n,
    lambda n: ['"abc"'] + ["2"] * (n - 1),
    lambda n: ["3", "2", "1"] + ["1"] * (n - 3) if n >= 3 else ["3"] * n,
    lambda n: ["2"] + ['"b"'] * (n - 1),
    lambda n: ['"FALSE"'] * n,
    lambda n: ['"0"'] * n,
    lambda n: ['""'] * n,
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
                    def call(x):
                        a = list(args)
                        a[p] = x
                        return f"{prefix}{nm}({','.join(a)})"
                    rows.append((ns, name, n, p, [call(MARKER)] + [call(c) for _, c in CANDIDATES]))
    seen, uniq = set(), []
    for r in rows:
        if r[4][0] not in seen:
            seen.add(r[4][0])
            uniq.append(r)
    rows = uniq
    allf = [f for r in rows for f in r[4]]
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


def norm(a):
    if a[0] == "boolean":
        return ("float", "1" if a[1] == "true" else "0")
    if a[0] in ("float", "percentage", "currency") and a[1] in ("-0",):
        return ("float", "0")
    return tuple(a)


def classify(raw):
    """`{(name, ns): {position: reading}}`: `number`, `text`, `empty`, `refused` or `other`."""
    res = {}
    k = 1 + len(CANDIDATES)
    for i, (ns, name, n, p, fs) in enumerate(raw["rows"]):
        if name in RANDOM:
            continue
        ans = raw["answers"][k * i : k * (i + 1)]
        m, cands = ans[0], ans[1:]
        if "crash" in [a[0] for a in ans]:
            continue
        # A call that is an error for every reading says nothing about the element.
        if m[0] == "error" and all(c[0] == "error" for c in cands) and m[1] != "Err:504":
            continue
        if m[0] == "error" and m[1] == "Err:504" and not any(norm(c) == norm(m) for c in cands):
            what = "refused"
        else:
            what = "other"
            for (kind, _), c in zip(CANDIDATES, cands):
                if norm(c) == norm(m):
                    what = "number" if kind in ("number", "bool") else kind
                    break
            else:
                what = "other"
        res.setdefault((name, ns), {}).setdefault(p, []).append(what)
    out = {}
    for key, pos in res.items():
        for p, whats in pos.items():
            # The most specific reading that any sample showed: a sample where the readings agree is
            # `number` (the first candidate), so a text or an empty or a refusal is what the function does.
            for kind in ("refused", "text", "empty", "other", "number"):
                if kind in whats:
                    out.setdefault(key, {})[p] = kind
                    break
    return out


def table(raw_path):
    raw = json.load(open(raw_path))
    cls = classify(raw)
    kinds = {k: [] for k in ("text", "empty", "refused")}
    other = []
    for (name, ns), pos in sorted(cls.items()):
        for kind in kinds:
            mask = sum(1 << p for p, w in pos.items() if w == kind)
            if mask:
                kinds[kind].append((name, ns, mask))
        for p, w in pos.items():
            if w == "other":
                other.append((name, ns, p))
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("//! How each function that takes one value reads the element that the missing `else` of an\n")
        f.write("//! `IF` gave, when it takes the array element by element, in `LibreOffice` 26.8: the number 0\n")
        f.write("//! unless the function is in a table here (the text `FALSE` or an empty text).\n//!\n")
        f.write("//! Generated by `sheetrs-compiler/scripts/probe_marker.py`: do not edit.\n\n")
        for kind, const in (("text", "TEXT"), ("empty", "EMPTY")):
            f.write(f"/// `(name, namespace, arguments)`, sorted by name and namespace: a bit per argument, the first is bit 0.\n")
            f.write(f"pub(super) const {const}: &[(&str, &str, u32)] = &[\n")
            for name, ns, mask in kinds[kind]:
                f.write(f'    ("{name}", "{ns}", {"u32::MAX" if mask == 0xFFFFFFFF else mask}),\n')
            f.write("];\n\n")
    print({k: len(v) for k, v in kinds.items()}, "refused:", kinds["refused"], "other:", other, file=sys.stderr)


if __name__ == "__main__":
    {"ask": lambda: ask(sys.argv[2]), "table": lambda: table(sys.argv[2])}[sys.argv[1]]()

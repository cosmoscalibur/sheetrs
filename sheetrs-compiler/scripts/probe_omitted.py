#!/usr/bin/env python3
"""An argument left out (`TAKE(r,)`, `LOG(10,)`, `DATE(,1,1)`): what `LibreOffice` makes of it, for every
function of the registry and each of its positions, and the table the engine takes from it
(`sheetrs-engine/src/functions/omitted_table.rs`).

    scripts/probe_omitted.py ask RAW.json      # asks LibreOffice 26.8 (needs soffice, openpyxl, cargo)
    scripts/probe_omitted.py table RAW.json    # writes omitted_table.rs from the answers

An argument that is written empty is a value, not an absent one: `LibreOffice` takes it for the number 0
(`LOG(10,)` is the logarithm in base 0, `DATE(2020,,5)` is the fifth day of month 0, `TAKE(r,)` takes 0
rows, an invalid argument), for an empty text where the function wants text (`SUBSTITUTE("abc","b",)`)
or as if it was not there (`ROUND(1.5,)`, `MATCH(5,r,)` with its default type). Each call is asked with the
position empty, with 0, with an empty text and, for the last position, with the argument removed; the
position is the one of those that gives the same answer, and when none does and the answer is the same
error everywhere, the error.
"""
import collections
import json
import os
import sys
from concurrent.futures import ThreadPoolExecutor

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import lo_probe  # noqa: E402
import probe_arity as pa  # noqa: E402

ROOT = os.path.join(HERE, "..", "..")
OUT = os.path.join(ROOT, "sheetrs-engine", "src", "functions", "omitted_table.rs")
BASES = [
    lambda n: ["2"] * n,
    lambda n: ["Data!A1:A8"] + ["2"] * (n - 1),
    lambda n: ['"abc"'] + ["2"] * (n - 1),
    lambda n: ["2", "Data!A1:A8"] + ["2"] * (n - 2) if n >= 2 else ["2"] * n,
    lambda n: ["3", "2", "1"] + ["1"] * (n - 3) if n >= 3 else ["3"] * n,
]


def samples(funcs, spelling):
    out = []
    for ns, name, mn, mx in funcs:
        if (ns, name) not in spelling:
            continue
        prefix, nm = spelling[(ns, name)]
        for n in sorted({max(mn, 1), min(mx, max(mn, 1) + 1)}):
            if n > 5:
                continue
            for base in BASES:
                args = base(n)[:n]
                for p in range(n):
                    def call(sub, drop=False):
                        a = list(args)
                        if drop:
                            a = a[:p]
                        else:
                            a[p] = sub
                        return f"{prefix}{nm}({','.join(a)})"
                    out.append({
                        "ns": ns, "name": name, "n": n, "p": p,
                        "o": call(""), "z": call("0"), "e": call('""'),
                        "d": call("", drop=True) if p == n - 1 and n - 1 >= mn else None,
                    })
    seen, uniq = set(), []
    for s in out:
        if s["o"] not in seen:
            seen.add(s["o"])
            uniq.append(s)
    return uniq


def ask(out_path):
    funcs = [(ns, name, int(mn), int(mx)) for ns, name, mn, mx in pa.registry()]
    funcs = [f for f in funcs if (f[0], f[1]) not in pa.NOT_PROBED and not f[1].startswith("_")]
    spell_path = out_path + ".spelling"
    if os.path.exists(spell_path):
        spelling = {tuple(k.split("|")): tuple(v) for k, v in json.load(open(spell_path)).items()}
    else:
        spelling = pa.find_spellings(funcs)
        json.dump({"|".join(k): list(v) for k, v in spelling.items()}, open(spell_path, "w"))
    ss = samples(funcs, spelling)
    allf = sorted({x for s in ss for x in (s["o"], s["z"], s["e"], s["d"]) if x})
    print(len(ss), "samples,", len(allf), "formulas", file=sys.stderr)
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
    flat = {f: a for chunk, ans in zip(chunks, answers) for f, a in zip(chunk, ans)}
    json.dump({"samples": ss, "answers": flat}, open(out_path, "w"))


CODES = {"Err:502": "502", "Err:504": "504", "Err:511": "511", "Err:518": "518", "Err:519": "519", "#VALUE!": "v", "#N/A": "a", "#NAME?": "n", "#REF!": "r", "#DIV/0!": "d", "#NUM!": "u"}


def table(raw_path):
    raw = json.load(open(raw_path))
    ans = raw["answers"]
    votes = collections.defaultdict(lambda: collections.defaultdict(collections.Counter))
    for s in raw["samples"]:
        # One empty argument (`F()`) is no argument: the call has none.
        if s["n"] == 1:
            continue
        o, z, e = ans.get(s["o"]), ans.get(s["z"]), ans.get(s["e"])
        d = ans.get(s["d"]) if s["d"] else None
        if not o or not z or not e or "crash" in (o[0], z[0], e[0]):
            continue
        key = (s["name"], s["ns"])
        v = votes[key][s["p"]]
        eq = lambda a, b: a is not None and tuple(a) == tuple(b)
        same_d, same_z, same_e = eq(d, o), eq(z, o), eq(e, o)
        if same_d and not same_z and not same_e:
            v["D"] += 1
        elif same_z and not same_d and not same_e:
            v["Z"] += 1
        elif same_e and not same_z and not same_d:
            v["E"] += 1
        # 0 and an empty text are not told apart by this call, and the argument removed is not the same.
        elif not (same_d or same_z or same_e) and o[0] == "error" and o[1] in CODES:
            v["!" + CODES[o[1]]] += 1
    # `IFERROR(,x)` is `Err:518` and `IFNA(NA(),)` is `#N/A`: the functions answer them themselves.
    for key, positions in {("IFERROR", "core"): (0, 1), ("IFNA", "core"): (0, 1)}.items():
        for p in positions:
            votes[key].pop(p, None)
    entries = []
    for (name, ns), by_pos in sorted(votes.items()):
        parts = []
        for p in sorted(by_pos):
            c = by_pos[p]
            if not c:
                continue
            best, n = c.most_common(1)[0]
            # A position that is sometimes one thing and sometimes another is left to the default.
            if n < sum(c.values()) / 2:
                continue
            parts.append((p, best))
        if parts:
            entries.append((name, ns, parts))
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("//! What `LibreOffice` 26.8 makes of an argument written empty (`LOG(10,)`, `DATE(,1,1)`), by function.\n//!\n")
        f.write("//! Generated by `sheetrs-compiler/scripts/probe_omitted.py`: do not edit.\n\n")
        f.write("/// `(name, namespace, [(argument, what)])`, sorted by name and namespace: the argument (from 0) is\n")
        f.write("/// `Z` the number 0, `E` an empty text, `D` as if it was not written (the last argument), or `!code` the\n")
        f.write("/// error `Err:code` (`!v` is `#VALUE!`); an argument that is not listed is taken as absent.\n")
        f.write("pub(super) type Positions = &'static [(u8, &'static str)];\n\n")
        f.write("/// The table: see [`Positions`].\n")
        f.write("pub(super) const OMITTED: &[(&str, &str, Positions)] = &[\n")
        for name, ns, parts in entries:
            inner = ", ".join(f'({p}, "{w}")' for p, w in parts)
            f.write(f'    ("{name}", "{ns}", &[{inner}]),\n')
        f.write("];\n")
    print(len(entries), "functions written", file=sys.stderr)


if __name__ == "__main__":
    {"ask": lambda: ask(sys.argv[2]), "table": lambda: table(sys.argv[2])}[sys.argv[1]]()

#!/usr/bin/env python3
"""What every function of the registry does with a reference over several sheets (`Data!A1:Wide!A3`),
in `LibreOffice`, and the table the engine takes from it (`sheetrs-engine/src/functions/threed.rs`).

    scripts/probe_3d.py ask [RAW.json]        # asks LibreOffice 26.8 (needs soffice, openpyxl, cargo)
    scripts/probe_3d.py controls RAW.json     # the same calls over one cell, to tell what is the reference's
    scripts/probe_3d.py table RAW.json        # writes threed_table.rs from the answers

A reference from one sheet to another is read by `LibreOffice` as the same area on every sheet from
the first to the last (the sheets in between included, in the order of the workbook): `Data!A1:Wide!A3`,
and `Data:Wide!A1:A3`, which is how an XLSX writes it. A function either goes through the cells of all the
sheets (`SUM`, `COUNT`, `MAX`, `ROWS`…) or refuses it: with `Err:504` (`SUMIF`, `VLOOKUP`, `MATCH`,
`SUMPRODUCT`…), with `#REF!` (`INDEX`) or with `#VALUE!` (where the function takes one value, and in an
operation). The calls are those of `probe_intersection.py` (a reference in each position of the arguments
the function takes) with the reference over two sheets; a position is refused with the error it
answers when the same call over one sheet of the area does not answer it.
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
OUT = os.path.join(ROOT, "sheetrs-engine", "src", "functions", "threed.rs")
KINDS = [
    ("Data!A1:Wide!A3", "Data!A1:A3"),
    ("Data!A1:Wide!C3", "Data!A1:C3"),
    ("Data!A1:Wide!C1", "Data!A1:C1"),
]
SHAPES = [
    ["{R}"], ["{R}", "1"], ["1", "{R}"], ["{R}", "{R}"], ["{R}", "{R}", "1"], ["1", "{R}", "{R}"],
    ["{R}", "1", "{R}"], ["1", "1", "{R}"], ["1", "{R}", "1"], ["{R}", "1", "1"], ["{R}", "1", "1", "1"],
    ['"a"', "{R}"], ["{R}", '"a"'], ["{R}", '">0"'], ["{R}", "{R}", "{R}"],
]


def formulas(funcs, spelling):
    rows = []
    for ns, name, mn, mx in funcs:
        if (ns, name) not in spelling:
            continue
        prefix, nm = spelling[(ns, name)]
        for shape in SHAPES:
            n = len(shape)
            if n > mx:
                continue
            if n < mn:
                shape = shape + ["1"] * (mn - n)
            for three, one in KINDS:
                for expr in (False, True):
                    args3 = [(f"{three}*1" if expr else three) if a == "{R}" else a for a in shape]
                    args1 = [(f"{one}*1" if expr else one) if a == "{R}" else a for a in shape]
                    rows.append((ns, name, f"{prefix}{nm}({','.join(args3)})", f"{prefix}{nm}({','.join(args1)})", [i for i, a in enumerate(shape) if a == "{R}"]))
    return rows


def ask(out_path):
    funcs = [(ns, name, int(mn), int(mx)) for ns, name, mn, mx in pa.registry()]
    funcs = [f for f in funcs if (f[0], f[1]) not in pa.NOT_PROBED and not f[1].startswith("_")]
    spell_path = out_path + ".spelling"
    if os.path.exists(spell_path):
        spelling = {tuple(k.split("|")): tuple(v) for k, v in json.load(open(spell_path)).items()}
    else:
        spelling = pa.find_spellings(funcs)
        json.dump({"|".join(k): list(v) for k, v in spelling.items()}, open(spell_path, "w"))
    rows = formulas(funcs, spelling)
    allf = sorted({r[2] for r in rows} | {r[3] for r in rows})
    print(len(rows), "rows,", len(allf), "formulas", file=sys.stderr)
    chunks = [allf[i : i + 800] for i in range(0, len(allf), 800)]

    def run(chunk):
        masked = chunk
        try:
            return lo_probe.probe(masked)
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
    json.dump({"rows": rows, "answers": flat}, open(out_path, "w"))


CODES = {"Err:504": "4", "Err:502": "2", "#VALUE!": "v", "#REF!": "r", "#NAME?": "n", "#N/A": "a"}
THREE = "Data!A1:Wide!A3"
CELL = "Data!A1"


def controls(raw_path):
    """The same calls over one cell (`Data!A1`) in place of the reference over sheets: an error of the
    reference that the call over a cell does not give is the function refusing the reference."""
    raw = json.load(open(raw_path))
    cells = sorted({r[2].replace(THREE, CELL) for r in raw["rows"] if r[2].count(THREE) == 1 and "*1" not in r[2]})
    print(len(cells), "controls", file=sys.stderr)
    chunks = [cells[i : i + 800] for i in range(0, len(cells), 800)]

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
    json.dump(flat, open(raw_path + ".controls", "w"))


def table(raw_path):
    raw = json.load(open(raw_path))
    ctl = json.load(open(raw_path + ".controls"))
    funcs = {(n, ns): mx for ns, n, mn, mx in pa.registry()}
    seen = collections.defaultdict(lambda: collections.defaultdict(lambda: {"clean": 0, "refused": collections.Counter()}))
    for ns, name, f3, f1, pos in raw["rows"]:
        if len(pos) != 1 or f3.count(THREE) != 1:
            continue
        a3 = raw["answers"].get(f3)
        c = ctl.get(f3.replace(THREE, CELL))
        if not a3 or not c or "crash" in (a3[0], c[0]):
            continue
        bucket = seen[(name, ns)][pos[0]]
        if a3[0] != "error":
            bucket["clean"] += 1
        elif a3[1] in CODES and not (c[0] == "error" and c[1] == a3[1]):
            bucket["refused"][CODES[a3[1]]] += 1
    entries = []
    for (name, ns), by_pos in sorted(seen.items()):
        chars = []
        for p in range(max(by_pos) + 1):
            b = by_pos.get(p)
            if not b or b["clean"] or not b["refused"]:
                chars.append(".")
            else:
                chars.append(b["refused"].most_common(1)[0][0])
        if any(ch != "." for ch in chars):
            # The last character stands for every later argument: the others of a function with
            # any number of arguments take what the last one sampled does.
            entries.append((name, ns, "".join(chars)))
    entries.sort()
    with open(OUT.replace("threed.rs", "threed_table.rs"), "w", encoding="utf-8") as f:
        f.write("//! What `LibreOffice` 26.8 does with a reference over several sheets in each argument of each function.\n//!\n")
        f.write("//! Generated by `sheetrs-compiler/scripts/probe_3d.py`: do not edit. A string per function, a character per\n")
        f.write("//! argument from 0: `.` the function goes through the cells of the sheets, `4`, `2`, `v`, `r`, `n` or `a` it\n")
        f.write("//! refuses the reference with `Err:504`, `Err:502`, `#VALUE!`, `#REF!`, `#NAME?` or `#N/A`; the last\n")
        f.write("//! character stands for every later argument.\n\n")
        f.write("/// `(name, namespace, outcome of each argument)`, sorted by name and namespace.\n")
        f.write("pub(super) const THREED: &[(&str, &str, &str)] = &[\n")
        for name, ns, string in entries:
            f.write(f'    ("{name}", "{ns}", "{string}"),\n')
        f.write("];\n")
    print(len(entries), "functions written", file=sys.stderr)


if __name__ == "__main__":
    {"ask": lambda: ask(sys.argv[2]), "controls": lambda: controls(sys.argv[2]), "table": lambda: table(sys.argv[2])}[sys.argv[1]]()

#!/usr/bin/env python3
"""Ask LibreOffice what every function of the engine's registry does with each number of arguments, and
write `sheetrs-engine/src/functions/arity.rs`.

    scripts/probe_arity.py            # needs cargo, openpyxl and soffice ($SOFFICE: the 26.8 build)

For each function of the registry (`cargo run -p sheetrs-engine --example functions`), spelled the way an
XLSX writes it (`_xlfn.` and the like, found by trying: LibreOffice answers `#NAME?` to a spelling it does
not know), a call of 0 up to `max + 2` arguments is asked with several arguments (numbers, texts, ranges, cells, a range
in each position that a function takes its table or series in). A count is **refused** when LibreOffice gives
an arity error (`Err:504` parameter list, `Err:508` parentheses, `Err:509` missing operator, `Err:510` and `Err:511`
missing variable, `Err:518` internal syntax) with every
one of them, and the commonest of those is its code; a value, or any other error, for some arguments
means the count is taken (the error is the function's own). The result is one string per
function, a character per count from 0: `.` the count is taken, `4`, `8` or `1` the code of the refusal;
the last character stands for every larger count.

The engine uses it (`functions::arity`) to answer a call that LibreOffice refuses with the very code
LibreOffice gives, and to know the counts it takes that the engine does not implement.
"""
import collections
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import lo_probe  # noqa: E402

ROOT = os.path.join(HERE, "..", "..")
OUT = os.path.join(ROOT, "sheetrs-engine", "src", "functions", "arity.rs")

# How an XLSX writes each namespace, in the order to try. Names LibreOffice has under another spelling
# (`FORECAST.ETS.ADD` is `_xlfn.FORECAST.ETS`) are tried without their `.ADD`.
PREFIXES = {
    "core": ["", "_xlfn.", "_xlfn.ORG.OPENOFFICE.", "_xlfn.ORG.LIBREOFFICE.", "_xlfn._xlws."],
    "microsoft": ["_xlfn.", "", "_xlfn._xlws.", "_xlfn.ORG.OPENOFFICE.", "_xlfn.ORG.LIBREOFFICE."],
    "openoffice": ["_xlfn.ORG.OPENOFFICE.", "", "_xlfn."],
    "libreoffice": ["_xlfn.ORG.LIBREOFFICE.", "_xlfn."],
    "legacy": ["", "_xlfn."],
}
# `SORT` and `FILTER` are `_xlfn._xlws.` in an XLSX (`_xlfn.SORT` is another function of LibreOffice).
PREFIXES_OF = {"SORT": ["_xlfn._xlws."], "FILTER": ["_xlfn._xlws."]}
# An XLSX cannot say `core::CEILING` (a bare `CEILING` there is Excel's), so these are not probed.
NOT_PROBED = {("core", "CEILING"), ("core", "FLOOR"), ("core", "FDIST"), ("core", "FINV")}
# The arguments of a call, by position: a call is refused only if every one of these gives the same
# arity error (an error that depends on what is given is the function's own). Uniform ones, and ones
# with a range where functions take their table or their series, and numbers elsewhere.
R = "Data!A1:A3"
PATTERNS = [
    lambda n: ["1"] * n,
    lambda n: ['"a"'] * n,
    lambda n: [R] * n,
    lambda n: ["Data!A1"] * n,
    lambda n: [R] + ["1"] * (n - 1),
    lambda n: ["1", R] + ["1"] * (n - 2) if n >= 2 else ["1"] * n,
    lambda n: [R, R] + ["1"] * (n - 2) if n >= 2 else [R] * n,
    lambda n: ["1", "1", R] + ["1"] * (n - 3) if n >= 3 else ["1"] * n,
    lambda n: [R, "1", R] + ["1"] * (n - 3) if n >= 3 else ["1"] * n,
    lambda n: ["1", R, R] + ["1"] * (n - 3) if n >= 3 else ["1"] * n,
    lambda n: [R, R, "1", R] + ["1"] * (n - 4) if n >= 4 else ["1"] * n,
    lambda n: ['"a"', R] + ['"a"'] * (n - 2) if n >= 2 else ['"a"'] * n,
]
DATABASE = ["DAVERAGE", "DCOUNT", "DCOUNTA", "DGET", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM", "DVAR", "DVARP"]
# A function LibreOffice refuses every count of with numbers, but that answers another error to some other
# kind of argument (so the classification below calls the count taken): `RANDARRAY` takes at most five
# arguments, a sixth is `Err:509` (checked with numbers, booleans and ranges).
OVERRIDES = {("microsoft", "RANDARRAY"): "......9"}
CODES = {"Err:504": "4", "Err:508": "8", "Err:509": "9", "Err:510": "0", "Err:511": "1", "Err:518": "b"}


def registry():
    out = subprocess.run(
        ["cargo", "run", "-q", "-p", "sheetrs-engine", "--example", "functions"],
        cwd=ROOT, capture_output=True, text=True, check=True,
    ).stdout
    return [line.split() for line in out.splitlines()]


def counts(minimum, maximum):
    top = minimum + 3 if maximum == 255 else min(maximum + 2, 34)
    return range(top + 1)


def call(prefix, name, pattern, n):
    return f"{prefix}{name}({','.join(pattern(n)[:n])})"


def find_spellings(funcs):
    """`{(namespace, name): (prefix, name)}`: how LibreOffice reads each function of `funcs` in an
    XLSX (it answers `#NAME?` to a spelling it does not know)."""
    spelling = {}
    todo = list(funcs)
    for step in range(max(len(v) for v in PREFIXES.values())):
        rows = []
        for ns, name, mn, mx in todo:
            names = [name] + ([name[:-4]] if name.endswith(".ADD") else [])
            prefixes = PREFIXES_OF.get(name, PREFIXES[ns])
            if step < len(prefixes):
                for nm in names:
                    for n in (mn, mn + 1):
                        rows.append((ns, name, prefixes[step], nm, call(prefixes[step], nm, PATTERNS[0], n)))
        answers = lo_probe.probe_all([r[4] for r in rows])
        for r, a in zip(rows, answers):
            if a[1] != "#NAME?" and a[0] != "crash" and (r[0], r[1]) not in spelling:
                spelling[(r[0], r[1])] = (r[2], r[3])
        todo = [f for f in todo if (f[0], f[1]) not in spelling]
    return spelling


def main():
    funcs = [(ns, name, int(mn), int(mx)) for ns, name, mn, mx in registry()]
    funcs = [f for f in funcs if (f[0], f[1]) not in NOT_PROBED and not f[1].startswith("_")]
    spelling = find_spellings(funcs)
    # 2. Every count with every kind of argument.
    rows = []
    for ns, name, mn, mx in funcs:
        if (ns, name) in spelling:
            prefix, nm = spelling[(ns, name)]
            for n in counts(mn, mx):
                for pattern in PATTERNS:
                    rows.append((ns, name, n, pattern, call(prefix, nm, pattern, n)))
    answers = lo_probe.probe_all([r[4] for r in rows])
    seen = collections.defaultdict(lambda: collections.defaultdict(list))
    for r, a in zip(rows, answers):
        if a[0] == "error" and a[1] in CODES:
            what = CODES[a[1]]
        elif a[0] == "error" and a[1] == "#VALUE!" and R in r[4]:
            # A range handed to a plain argument of a formula that is not an array formula is
            # intersected with the row of the formula: a `#VALUE!` that says nothing of the count.
            what = "?"
        else:
            what = "."
        seen[(r[0], r[1])][r[2]].append(what)
    table = []
    for ns, name, mn, mx in funcs:
        got = seen.get((ns, name))
        if not got:
            continue
        string = ""
        for n in sorted(got):
            kinds = got[n]
            refusals = [k for k in kinds if k not in ".?"]
            # Inside the counts the engine takes, a count is refused only when every one of the
            # arguments is refused (a call the engine runs must not be refused on a doubt); outside
            # them, unless something answered a value or an error of its own (a `#VALUE!` that a range
            # got from the intersection with the row says nothing).
            strict = mn <= n <= mx
            refused = refusals and "." not in kinds and (not strict or "?" not in kinds)
            string += collections.Counter(refusals).most_common(1)[0][0] if refused else "."
        # A function every count of which looks refused answers an error of its own to every kind of
        # argument we gave it (the database functions need a table); its counts are the engine's.
        if "." not in string:
            if name in DATABASE:
                string = "".join("." if n == 3 else "4" for n in range(len(string)))
            elif name in ("CHOOSECOLS", "CHOOSEROWS"):
                string = "11" + "." * (len(string) - 2)
            else:
                continue
        table.append((name, ns, OVERRIDES.get((ns, name), string)))
    table.sort()
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("//! What `LibreOffice` 26.8 does with each number of arguments of each function of the registry.\n")
        f.write("//!\n//! Generated by `sheetrs-compiler/scripts/probe_arity.py`: do not edit. A string per function, a character\n")
        f.write("//! per number of arguments from 0: `.` the count is taken; `4`, `8`, `9`, `0`, `1` or `b` it is refused with\n")
        f.write("//! `Err:504`, `Err:508`, `Err:509`, `Err:510`, `Err:511` or `Err:518`; the last character stands for every\n")
        f.write("//! larger count.\n\n")
        f.write("/// `(name, namespace, counts)`, sorted by name and namespace.\n")
        f.write("pub(super) const ARITY: &[(&str, &str, &str)] = &[\n")
        for name, ns, string in table:
            f.write(f'    ("{name}", "{ns}", "{string}"),\n')
        f.write("];\n")
    print(f"{len(table)} functions written to {os.path.relpath(OUT, ROOT)}", file=sys.stderr)
    missing = [(f[0], f[1]) for f in funcs if (f[0], f[1]) not in {(t[1], t[0]) for t in table}]
    print("not in the table:", missing, file=sys.stderr)


if __name__ == "__main__":
    main()

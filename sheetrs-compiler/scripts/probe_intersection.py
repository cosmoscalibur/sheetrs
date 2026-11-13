#!/usr/bin/env python3
"""Where does a range meet the formula's own cell? Sweep every function of the registry with ranges and
range expressions in each argument position, in LibreOffice and in the engine, and list the functions
whose engine disagrees with LibreOffice on that.

    scripts/probe_intersection.py [OUT.json]       # needs cargo, openpyxl and soffice ($SOFFICE: 26.8)

In a formula that is not an array formula a range given where a function takes a value is met by the
row (a column range), the column (a row range) or both (a block) of the formula's cell, in any sheet,
and is `#VALUE!` when they do not meet; an operator over a range does the same, and so does a function
that takes no matrix there. Only the arguments `LibreOffice` evaluates as arrays (`SUMPRODUCT`, `MMULT`,
`CORREL`…) see the whole range. Each function is called with the `Wide` sheet of the semantics fixture
(`gen_semantics_fixture.build`) as a plain range and as the expression `range*1`, with a count of
arguments the engine takes, and each call is paired with the same call over the one cell that the
range is met in (`Wide!A17` for a range in row 17). A call is *met* when its answer is the one of its
pair; the engine must be met where LibreOffice is, and not met where it is not. What a function
answers to its own arguments (the error of a domain, the format of a date) is not under test: it is
the same for the pair.
"""
import collections
import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import gen_semantics_fixture as gen  # noqa: E402
import lo_probe  # noqa: E402
import probe_arity as pa  # noqa: E402

ROOT = os.path.join(HERE, "..", "..")
SHEETC = os.environ.get("SHEETC", os.path.join(ROOT, "target", "release", "sheetc"))
ROWS = 40000
COL = f"Wide!A1:A{ROWS}"
TXT = f"Wide!B1:B{ROWS}"
NUM = f"Wide!C1:C{ROWS}"
BLOCK = f"Wide!A1:C{ROWS}"
ROW = "Wide!A1:C1"
# Argument shapes by position: `{R}` is replaced by a plain range or by `range*1`.
SHAPES = [
    ["{R}"], ["{R}", "1"], ["1", "{R}"], ["{R}", "{R}"], ["{R}", "{R}", "1"], ["1", "{R}", "{R}"],
    ["{R}", "1", "{R}"], ["1", "1", "{R}"], ["{R}", "{R}", "{R}"], ['"a"', "{R}"], ["{R}", '"a"'],
]
KINDS = [COL, TXT, BLOCK, ROW, NUM]


def formulas_for(funcs, spelling):
    rows = []
    for ns, name, mn, mx in funcs:
        if (ns, name) not in spelling:
            continue
        prefix, nm = spelling[(ns, name)]
        for shape in SHAPES:
            n = len(shape)
            if not mn <= n <= mx:
                # pad with numbers up to the smallest count the function takes
                if n < mn:
                    shape = shape + ["1"] * (mn - n)
                    n = mn
                if n > mx:
                    continue
            for expr in (False, True):
                for rng in KINDS:
                    args = [(f"{rng}*1" if expr else rng) if a == "{R}" else a for a in shape]
                    rows.append((ns, name, f"{prefix}{nm}({','.join(args)})"))
    return rows


def efmt(v):
    if isinstance(v, dict):
        return v.get("code") or v["error"]
    if isinstance(v, bool):
        return "true" if v else "false"
    if v is None:
        return ""
    return v


def same(a, b):
    try:
        x, y = float(a), float(b)
        return abs(x - y) <= 1e-9 * max(1, abs(x), abs(y))
    except (TypeError, ValueError):
        return str(a) == str(b)


CHUNK = 1000


def engine_answers(chunk):
    """What the engine answers for the formulas of one workbook (one row each), or `None` when it fails."""
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "sweep.xlsx")
        gen.build(path, chunk)
        res = subprocess.run([SHEETC, "calculate", "-i", path], input="", capture_output=True, text=True)
    if "{" not in res.stdout:
        return None
    answer = json.loads(res.stdout[res.stdout.index("{"):])["valueRanges"]
    rows = next(e["values"] for e in answer if e["name"].startswith("Calc!"))
    return [efmt(row[0]) for row in rows]


def safe_probe(chunk, lo=None, hi=None):
    """`lo_probe.probe` of the rows `lo:hi` of `chunk` with the others blanked (so every formula keeps
    its row, which the answer depends on); a formula that makes LibreOffice fail is answered `crash`."""
    lo = 0 if lo is None else lo
    hi = len(chunk) if hi is None else hi
    masked = [f if lo <= i < hi else "0" for i, f in enumerate(chunk)]
    try:
        return lo_probe.probe(masked)[lo:hi]
    except subprocess.CalledProcessError:
        if hi - lo == 1:
            return [("crash", "")]
        mid = (lo + hi) // 2
        return safe_probe(chunk, lo, mid) + safe_probe(chunk, mid, hi)


def at_cell(formula, row):
    """The formula with each range of `Wide` replaced by the cell where the formula of `row` meets it."""
    for rng, cell in (
        (BLOCK, f"Wide!A{row}"), (COL, f"Wide!A{row}"), (TXT, f"Wide!B{row}"), (NUM, f"Wide!C{row}"),
        (ROW, "Wide!A1"),
    ):
        formula = formula.replace(rng, cell)
    return formula


def libre_answers(index, chunk, cache_dir, tag):
    cache = os.path.join(cache_dir, f"{tag}{index}.json") if cache_dir else None
    if cache and os.path.exists(cache):
        return json.load(open(cache))
    answers = safe_probe(chunk)
    if cache:
        json.dump(answers, open(cache, "w"))
    return answers


def met(a, b):
    """Whether two answers (kind, value) are the same answer."""
    return a[0] == b[0] and same(a[1], b[1])


def sweep_chunk(index, chunk, cache_dir):
    """The calls of one workbook, with the pair of each over its cell, as LibreOffice and as the engine answer them."""
    pairs = [at_cell(f, i + 1) for i, f in enumerate(chunk)]
    lo_f = libre_answers(index, chunk, cache_dir, "lo")
    lo_g = libre_answers(index, pairs, cache_dir, "lc")
    en_f = engine_answers(chunk)
    en_g = engine_answers(pairs)
    out = []
    for i, f in enumerate(chunk):
        if en_f is None or en_g is None:
            out.append((index * CHUNK + i + 1, f, "engine failed", "", ""))
            continue
        if "crash" in (lo_f[i][0], lo_g[i][0]):
            continue
        lo_met = met(lo_f[i], lo_g[i])
        en_met = same(en_f[i], en_g[i])
        if lo_met != en_met:
            out.append((index * CHUNK + i + 1, f, "met" if lo_met else "not met", f"{lo_f[i][1]!r} / cell {lo_g[i][1]!r}", f"{en_f[i]!r} / cell {en_g[i]!r}"))
    return out


def main():
    from concurrent.futures import ThreadPoolExecutor

    out_path = sys.argv[1] if len(sys.argv) > 1 else None
    cache_dir = os.environ.get("CACHE")
    if cache_dir:
        os.makedirs(cache_dir, exist_ok=True)
    funcs = [(ns, name, int(mn), int(mx)) for ns, name, mn, mx in pa.registry()]
    funcs = [f for f in funcs if (f[0], f[1]) not in pa.NOT_PROBED and not f[1].startswith("_")]
    spelling = pa.find_spellings(funcs)
    rows = formulas_for(funcs, spelling)
    seen, uniq = set(), []
    for r in rows:
        if r[2] not in seen:
            seen.add(r[2])
            uniq.append(r)
    rows = uniq
    print(f"{len(rows)} formulas", file=sys.stderr)
    formulas = [r[2] for r in rows]
    chunks = [formulas[i : i + CHUNK] for i in range(0, len(formulas), CHUNK)]
    names = {r[2]: r[1] for r in rows}
    with ThreadPoolExecutor(4) as pool:
        results = list(pool.map(lambda ic: sweep_chunk(ic[0], ic[1], cache_dir), enumerate(chunks)))
    diffs = collections.defaultdict(list)
    for chunk_result in results:
        for row, f, why, lv, ev in chunk_result:
            diffs[names[f]].append([row, f, why, lv, ev])
    if out_path:
        json.dump(diffs, open(out_path, "w"), indent=1)
    for name, items in sorted(diffs.items()):
        print(f"{name}: {len(items)} differ, e.g. {items[0][1]}: LibreOffice {items[0][2]} ({items[0][3]}), engine {items[0][4]}")
    print(f"{len(diffs)} functions differ of {len(funcs)}", file=sys.stderr)


if __name__ == "__main__":
    main()

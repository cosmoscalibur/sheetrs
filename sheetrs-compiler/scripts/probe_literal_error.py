#!/usr/bin/env python3
"""Which positions of which functions turn an error written in the formula into `Err:504`?

    scripts/probe_literal_error.py [--write]  # needs cargo, openpyxl and soffice ($SOFFICE: the 26.8 build)
    scripts/probe_literal_error.py --strict [NAME,NAME…]  # the engine against LibreOffice at every position (`$SHEETC`)

An error token written in a formula (`#NAME?`, `#NUM!`, `#DIV/0!`, `#N/A`, `#NULL!`, `#VALUE!`) is not the
error a calculation gave. Where a function wants a reference (`OFFSET(#NAME?,0,0)`) LibreOffice finds no
reference in the token and answers `Err:504` (parameter list); the very same call with a calculated error
(`OFFSET(NA(),0,0)`) answers that error (`#N/A`), and where a function wants a value both propagate
(`OFFSET(Data!A1,#NAME?,0)` is `#NAME?`). A bare `#REF!` in an XLSX is a name that does not exist
(`#NAME?`), not an error token, so it is not probed.

For each function of the engine's registry (spelled as `probe_arity` finds it), each number of arguments
from the fewest it takes to two more, and each pattern of arguments (`probe_arity.PATTERNS`: numbers, texts,
ranges, cells, a range where functions take a table or a series), the call is made with the pattern as it
is: when LibreOffice takes it (a value, or an error of the calculation), a position of it is replaced by `#NAME?` and
by `NA()`. A position is a **reference position** when the token answers `Err:504` (`Err:502`, `#VALUE!` or `Err:517` for
the functions that want an array: `CORREL(#NAME?,Data!A1:A3)` is `Err:502`) and `NA()` does not (it answers
its own `#N/A`): only calls that are good by themselves count, so an `Err:504` that another argument
causes is never taken for the position's, and the answers that the call itself could give (`#VALUE!`)
only when the call gives a value. The other tokens are then asked of each hit: they answer the same.

With `--write` the table `sheetrs-engine/src/functions/literal_error_table.rs` is written: for each
function, the positions (from 0; the last one stands for the ones after it when the function takes any
number of arguments).

Which error wins when a call has both a written error and a calculated one is asked too (`--write` writes
`computed_error_table.rs`): in each good call with two or more arguments, a written error `#NAME?` goes to a
position and a calculated `NA()` to another, in both orders. A function whose answer is always the calculated
`#N/A` (`SUM(#NAME?,NA())` is `#N/A`, the written one does not win by being first) is in the table; a
function that gives something else with some pair (`IFERROR(1/0,#NAME?)` is `#NAME?`: it catches the
calculated one) is not.
"""
import collections
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import lo_probe  # noqa: E402
import probe_arity as pa  # noqa: E402

OTHER_TOKENS = ["#NUM!", "#DIV/0!", "#N/A", "#NULL!", "#VALUE!"]
# What a written error is where a function wants a reference or an array: the answer, as a character of the table.
CODES = {"Err:504": "4", "Err:502": "2", "#VALUE!": "v", "Err:517": "7"}
inverse_codes = {v: k for k, v in CODES.items()}
ROWS = os.path.join(HERE, "literal_error_formulas.txt")
# Rows of the oracle the table does not give: what a position that wants a value does with the token, what a
# calculated error does, and which error wins when a call has several (the first error a calculation gave, then
# the first error written where a value is wanted, then `Err:504`).
EXTRA = """
OFFSET(NA(),0,0)
OFFSET(1/0,0,0)
OFFSET(Data!A1,#NAME?,0)
OFFSET(Data!A1,0,#NUM!)
OFFSET(Data!A1,0,0,#NAME?)
OFFSET(Data!A1,0,0,1,#DIV/0!)
OFFSET(Data!A1,NA(),0)
OFFSET(Data!A1,1/0,0)
OFFSET(#NAME?,NA(),0)
OFFSET(NA(),#NAME?,0)
OFFSET(#NAME?,#NUM!,0)
OFFSET(#NUM!,#NAME?,0)
OFFSET(#NAME?,0,NA())
OFFSET(Data!A1,#NAME?,#NUM!)
OFFSET(Data!A1,NA(),1/0)
OFFSET(1/0,NA(),0)
OFFSET(NA(),1/0,0)
SUM(OFFSET(#NAME?,0,0))
OFFSET(#NAME?,0,0)+1
IFERROR(OFFSET(#NAME?,0,0),7)
ISERROR(OFFSET(#NAME?,0,0))
ERROR.TYPE(OFFSET(#NAME?,0,0))
VLOOKUP(#NAME?,Data!A1:B3,2,0)
VLOOKUP(#NAME?,Data!A1:B3,#NUM!,0)
VLOOKUP(#NAME?,Data!A1:B3,2,#NUM!)
VLOOKUP(#NAME?,Data!A1:B3,1/0,0)
VLOOKUP(1/0,Data!A1:B3,2,0)
VLOOKUP(1,#NAME?,2,0)
VLOOKUP(1,#NAME?,#NUM!,0)
VLOOKUP(1,#NAME?,1/0,0)
VLOOKUP(1,Data!A1:B3,#NAME?,0)
MATCH(#NAME?,Data!A1:A3,0)
MATCH(1,#NAME?,0)
MATCH(#NAME?,#NUM!,0)
MATCH(#NAME?,#NUM!,#DIV/0!)
MATCH(#NAME?,Data!A1:A3,#NUM!)
MATCH(#NAME?,Data!A1:A3,1/0)
MATCH(1/0,#NAME?,0)
MATCH(NA(),#NUM!,0)
SUMIF(#NAME?,1)
SUMIF(#NAME?,1,Data!A1:A3)
SUMIF(Data!A1:A3,1,#NAME?)
SUMIF(Data!A1:A3,#NAME?)
SUMIF(Data!A1:A3,1/0,#NAME?)
SUMIF(#NAME?,1/0)
COUNTIF(#NAME?,1)
COUNTBLANK(#NAME?)
COUNTBLANK(NA())
ROW(#NAME?)
ROW(NA())
COLUMN(#NAME?)
ROWS(#NAME?)
COLUMNS(#NAME?)
AREAS(#NAME?)
INDEX(#NAME?,1,1)
INDEX(NA(),1,1)
INDEX(Data!A1:A3,#NAME?,1)
INDIRECT(#NAME?)
INDIRECT(NA())
NPV(1,#NAME?)
NPV(#NAME?,1)
NPV(#NAME?,#NUM!)
NPV(0.1,Data!A1:A3,#NAME?)
SUM(#NAME?)
SUM(1,#NAME?)
SUM(#NAME?,NA())
SUMPRODUCT(#NAME?)
ERROR.TYPE(OFFSET(#NAME?,0,0))
ERROR.TYPE(SQRT(-1))
ERROR.TYPE(1/0)
COUNTIF(#NAME?,1)
COUNTIF(NA(),1)
COUNTBLANK(NA())
SUM(NA(),#NAME?)
SUM(1/0,#NAME?)
SUM(#NUM!,#NAME?)
SUM(#NAME?,#NUM!)
_xlfn.AGGREGATE(NA(),1,1)
_xlfn.AGGREGATE(1,NA(),1)
_xlfn.AGGREGATE(1,1,NA())
_xlfn.XLOOKUP(1,Data!A1:A3,NA())
_xlfn.XLOOKUP(1,NA(),Data!A1:A3)
_xlfn.XLOOKUP(NA(),Data!A1:A3,Data!A1:A3)
_xlfn.XLOOKUP(1,Data!A1:A3,Data!A1:A3,5,NA())
_xlfn.XLOOKUP(Data!A1,Data!A1:A3,Data!A1:A3,5,1/0)
_xlfn.XLOOKUP(1,Data!A1:A3,Data!A1:A3,5,0,NA())
_xlfn.SWITCH(1,1/0,5)
_xlfn.SWITCH(1,2,5,1/0,6)
_xlfn.SWITCH(1,#NAME?,5)
_xlfn.SWITCH(1,Data!A1:A3,5)
_xlfn.SWITCH(1,2,5,Data!A1:A3,6)
_xlfn.SWITCH(1,NA(),5,1/0,6)
_xlfn.SWITCH(1,1/0,5,NA(),6)
_xlfn.SWITCH(1,1/0,5,2,6)
_xlfn.SWITCH(1,1/0,5,1,6)
_xlfn.SWITCH(1,2,5,1/0,6,7)
ACCRINT("a","a","a","a","a",#NAME?,NA())
VDB(1,1,1,1,#NAME?,NA())
# a bare #REF!: an XLSX reads a name (#NAME?), an ODS the error
OFFSET(#REF!,0,0)
OFFSET(#REF!,1,1)
SUM(OFFSET(#REF!,0,0))
IFERROR(OFFSET(#REF!,0,0),7)
OFFSET(Data!A1,#REF!,0)
OFFSET(Data!A1,0,#REF!)
ROW(#REF!)
SUMIF(#REF!,1)
VLOOKUP(#REF!,Data!A1:B3,2,0)
VLOOKUP(1,#REF!,2,0)
MATCH(1,#REF!,0)
SUM(#REF!)
SUM(1,#REF!)
#REF!
#REF!+1
-#REF!
#REF!&"a"
ISERROR(#REF!)
ISREF(#REF!)
ERROR.TYPE(#REF!)
IFERROR(#REF!,7)
{#REF!}
{1,#REF!}
{1;#REF!}
SUM({#REF!})
SUM({1,#REF!})
Data!#REF!
Data!#REF!+1
SUM(Data!#REF!)
OFFSET(Data!#REF!,0,0)
SUMPRODUCT(#REF!,1)
SUMPRODUCT("a",#REF!)
_xlfn.HSTACK(1,#REF!)
"""


def extra_rows():
    return [line for line in EXTRA.strip().splitlines() if line and not line.startswith("# ")]
OUT = os.path.join(HERE, "..", "..", "sheetrs-engine", "src", "functions", "literal_error_table.rs")


def main():
    if "--strict" in sys.argv:
        i = sys.argv.index("--strict")
        names = set(sys.argv[i + 1].split(",")) if len(sys.argv) > i + 1 else set()
        return strict(names)
    funcs = [(ns, name, int(mn), int(mx)) for ns, name, mn, mx in pa.registry() if not name.startswith("_")]
    funcs = [f for f in funcs if (f[0], f[1]) not in pa.NOT_PROBED]
    spelling = pa.find_spellings(funcs)
    bases = []  # (ns, name, n, pattern index, [args], prefix+name)
    for ns, name, mn, mx in funcs:
        if (ns, name) not in spelling:
            continue
        prefix, nm = spelling[(ns, name)]
        top = min(mx, max(mn, 1) + 2)
        for n in range(max(mn, 1), top + 1):
            for k, pattern in enumerate(pa.PATTERNS):
                bases.append((ns, name, n, k, pattern(n)[:n], f"{prefix}{nm}"))
    answers = lo_probe.probe_all([f"{b[5]}({','.join(b[4])})" for b in bases])
    # A call is good by itself when LibreOffice takes it: a value, or an error of the calculation (`#N/A`
    # for a lookup that finds nothing) but not a refusal of the call (`Err:5xx`, or `#NAME?`).
    good = [
        b
        for b, a in zip(bases, answers)
        if a[0] not in ("crash", "empty") and not (a[0] == "error" and (a[1].startswith("Err:5") or a[1] == "#NAME?"))
    ]
    print(f"{len(good)} calls of {len(bases)} are good by themselves", file=sys.stderr)
    value_calls = {b[:4] for b, a in zip(bases, answers) if a[0] not in ("error", "crash", "empty")}
    rows = []
    for b in good:
        for i in range(len(b[4])):
            for arg in ("#NAME?", "NA()"):
                args = list(b[4])
                args[i] = arg
                rows.append((b, i, arg, f"{b[5]}({','.join(args)})"))
    answers = lo_probe.probe_all([r[3] for r in rows])
    got = {}
    for r, a in zip(rows, answers):
        got[(r[0][:4], r[1], r[2])] = a[1] if a[0] == "error" else "value"
    hits = {}
    for (key, i, arg), answer in got.items():
        calc = got.get((key, i, "NA()"))
        if arg != "#NAME?" or answer not in CODES or calc in (answer, "value", None):
            continue
        # `Err:504` is a refusal of the parameters whatever the call is; the others can be the call's own
        # error, so only a call that gives a value counts.
        if answer == "Err:504" or key in value_calls:
            hits[(key, i)] = CODES[answer]
    # `Err:502` and `#VALUE!` depend on the other arguments (`CORREL(#NAME?,1)` is `#NAME?`, `CORREL(#NAME?,Data!A1:A3)` is
    # `Err:502`): only where every call that gives a value answers it.
    answers_at = collections.defaultdict(set)
    for (key, i, arg), answer in got.items():
        if arg == "#NAME?" and key in value_calls:
            answers_at[(key[:2], i)].add(answer)
    for (key, i), code in list(hits.items()):
        if code != "4" and answers_at[(key[:2], i)] != {inverse_codes[code]}:
            del hits[(key, i)]
    # the other tokens of each hit
    check = []
    by_key = {(b[:4]): b for b in good}
    for (key, i) in hits:
        b = by_key[key]
        for t in OTHER_TOKENS:
            args = list(b[4])
            args[i] = t
            check.append(((key, i, t), f"{b[5]}({','.join(args)})"))
    results = lo_probe.probe_all([c[1] for c in check])
    inverse = {v: k for k, v in CODES.items()}
    for (tag, a) in zip((c[0] for c in check), results):
        if a[1] != inverse[hits[(tag[0], tag[1])]]:
            print(f"token differs: {tag} -> {a}", file=sys.stderr)
    if "--write" in sys.argv:
        write_computed_first(good)
        write_propagating(good)
    seen = collections.defaultdict(collections.Counter)
    lines = []
    for (key, i) in sorted(hits):
        ns, name, n, k = key
        seen[(ns, name, i)][hits[(key, i)]] += 1
        lines.append(f"{name}\targument {i + 1} of {n}\tpattern {k}\t{hits[(key, i)]}\t{' '.join(by_key[key][4])}")
    table = collections.defaultdict(dict)
    for (ns, name, i), codes in seen.items():
        if len(codes) > 1:
            print(f"{name} position {i + 1} answers {dict(codes)}: the commonest is taken", file=sys.stderr)
        table[(ns, name)][i] = codes.most_common(1)[0][0]
    print("\n".join(lines))
    print(f"\n{len(table)} functions", file=sys.stderr)
    if "--write" in sys.argv:
        write_table(table, {(f[0], f[1]): f[3] for f in funcs})
        write_rows(hits, by_key)
    if "--computed" in sys.argv:
        write_computed(hits, by_key, sys.argv[sys.argv.index("--computed") + 1])


def strict(names):
    """`--strict [NAME,NAME…]`: what the engine answers to every position of every good call with an error written there
    (`#NAME?`), next to what LibreOffice answers (the sweep that found the functions of `TOKEN_AS_ARRAY` and the rules of
    `error_where_a_reference_is_wanted`). Prints one line per difference and the count per function; needs `$SHEETC`
    (by default `target/release/sheetc`)."""
    import probe_intersection as pi

    funcs = [(ns, name, int(mn), int(mx)) for ns, name, mn, mx in pa.registry() if not name.startswith("_")]
    funcs = [f for f in funcs if (f[0], f[1]) not in pa.NOT_PROBED and (not names or f[1] in names)]
    spelling = pa.find_spellings(funcs)
    bases = []
    for ns, name, mn, mx in funcs:
        if (ns, name) not in spelling:
            continue
        prefix, nm = spelling[(ns, name)]
        for n in range(max(mn, 1), min(mx, max(mn, 1) + 2) + 1):
            for pattern in pa.PATTERNS:
                bases.append((name, f"{prefix}{nm}", pattern(n)[:n]))
    answers = lo_probe.probe_all([f"{b[1]}({','.join(b[2])})" for b in bases])
    good = [
        b for b, a in zip(bases, answers)
        if a[0] not in ("crash", "empty") and not (a[0] == "error" and (a[1].startswith("Err:5") or a[1] == "#NAME?"))
    ]
    rows = []
    for name, spelled, args in good:
        for i in range(len(args)):
            call = list(args)
            call[i] = "#NAME?"
            rows.append((name, f"{spelled}({','.join(call)})"))
    formulas = [r[1] for r in rows]
    libre = lo_probe.probe_all(formulas)
    engine = []
    for i in range(0, len(formulas), 1000):
        e = pi.engine_answers(formulas[i : i + 1000])
        engine += e if e is not None else [None] * len(formulas[i : i + 1000])
    differ = collections.Counter()
    for (name, f), l, e in zip(rows, libre, engine):
        same = e is not None and (str(e).lower() == str(l[1]).lower() if l[0] == "boolean" else pi.same(l[1], e))
        if not same:
            differ[name] += 1
            print(f"{f}\tLO={l[0]}:{l[1]}\tENG={e}")
    print(f"{sum(differ.values())} differ of {len(rows)}: {dict(differ)}", file=sys.stderr)


def write_propagating(good):
    """`propagating_error_table.rs`: the functions that answer an error a calculation gave in any argument,
    whatever the other arguments are (`ACCRINT("a","a","a","a","a","a",NA())` is `#N/A`, not the `#VALUE!`
    of the text): the error of the arguments is read before they are checked."""
    rows = []
    for b in good:
        for i in range(len(b[4])):
            args = list(b[4])
            args[i] = "NA()"
            rows.append((b[:2], f"{b[5]}({','.join(args)})"))
    answers = lo_probe.probe_all([r[1] for r in rows])
    verdict = {}
    for (key, _), a in zip(rows, answers):
        verdict[key] = verdict.get(key, True) and a == ("error", "#N/A")
    names = sorted((name, ns) for (ns, name), ok in verdict.items() if ok)
    lines = [
        "//! The functions that answer an error a calculation gave in any argument, as `LibreOffice` 26.8 does, whatever",
        "//! the other arguments are: the errors of the arguments are read before the arguments are checked.",
        "//!",
        "//! Generated by `sheetrs-compiler/scripts/probe_literal_error.py --write`; do not edit by hand.",
        "",
        "/// `(function, namespace)`.",
        "#[rustfmt::skip]",
        "pub(super) const PROPAGATING: &[(&str, &str)] = &[",
    ]
    lines += [f'    ("{n}", "{ns}"),' for n, ns in names]
    lines.append("];")
    with open(os.path.join(HERE, "..", "..", "sheetrs-engine", "src", "functions", "propagating_error_table.rs"), "w", encoding="utf-8") as out:
        out.write("\n".join(lines) + "\n")
    print(f"{len(names)} functions that answer the error of an argument", file=sys.stderr)


def write_computed_first(good):
    """`computed_error_table.rs`: the functions where a calculated error wins over a written one."""
    rows = []
    for b in good:
        n = len(b[4])
        if n < 2:
            continue
        # the first two and the first and the last positions, in both orders
        for i, j in {(0, 1), (1, 0), (0, n - 1), (n - 1, 0)}:
            args = list(b[4])
            args[i], args[j] = "#NAME?", "NA()"
            rows.append((b[:2], f"{b[5]}({','.join(args)})"))
    answers = lo_probe.probe_all([r[1] for r in rows])
    verdict = {}
    for (key, _), a in zip(rows, answers):
        verdict[key] = verdict.get(key, True) and a == ("error", "#N/A")
    names = sorted((name, ns) for (ns, name), ok in verdict.items() if ok)
    lines = [
        "//! The functions where an error a calculation gave wins over an error written in the formula, as `LibreOffice`",
        "//! 26.8 answers (`SUM(#NAME?,NA())` is `#N/A`), whatever the order of the arguments.",
        "//!",
        "//! Generated by `sheetrs-compiler/scripts/probe_literal_error.py --write`; do not edit by hand.",
        "",
        "/// `(function, namespace)`.",
        "#[rustfmt::skip]",
        "pub(super) const COMPUTED_FIRST: &[(&str, &str)] = &[",
    ]
    lines += [f'    ("{n}", "{ns}"),' for n, ns in names]
    lines.append("];")
    with open(os.path.join(HERE, "..", "..", "sheetrs-engine", "src", "functions", "computed_error_table.rs"), "w", encoding="utf-8") as out:
        out.write("\n".join(lines) + "\n")
    print(f"{len(names)} functions where a calculated error wins", file=sys.stderr)


def write_computed(hits, by_key, path):
    """The same calls with a calculated error (`NA()`, `1/0`) where the token was: LibreOffice answers that
    error; a check of the engine over every position of the table."""
    out, seen = [], set()
    for (key, i) in sorted(hits, key=lambda h: (h[0][1], h[1], h[0][2], h[0][3])):
        name = key[1]
        if (name, i) in seen:
            continue
        seen.add((name, i))
        b = by_key[key]
        for e in ("NA()", "1/0"):
            args = list(b[4])
            args[i] = e
            out.append(f"{b[5]}({','.join(args)})")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(out) + "\n")


def write_rows(hits, by_key):
    """The rows of the oracle: each position of the table with a good call, the token cycling through the six
    (all of them for the first position of each function), then the rows of `EXTRA`."""
    tokens = ["#NAME?"] + OTHER_TOKENS
    rows, seen, n = [], set(), 0
    for (key, i) in sorted(hits, key=lambda h: (h[0][1], h[1], h[0][2], h[0][3])):
        ns, name, count, k = key
        if (name, i) in seen:
            continue
        seen.add((name, i))
        b = by_key[key]
        first = not any(r[0] == name for r in rows)
        for t in tokens if first else [tokens[n % len(tokens)]]:
            args = list(b[4])
            args[i] = t
            rows.append((name, f"{b[5]}({','.join(args)})"))
        n += 1
    out = [r[1] for r in rows] + extra_rows()
    with open(ROWS, "w", encoding="utf-8") as f:
        f.write("\n".join(out) + "\n")


# What the sweep of `--strict` (every position of every good call, `scripts/probe_literal_error.py --strict`) found that the
# aggregation by position of `main` cannot say: in `CHOOSECOLS` and `CHOOSEROWS` only the array (the first argument) is a
# reference position, the lanes are values (`CHOOSECOLS(Data!A1:A3,#NAME?)` is `#NAME?`); and the arguments of the
# functions of `TOKEN_AS_ARRAY` (`sheetrs-engine/src/functions/literal_error.rs`) are read as arrays of one element
# that is the error, so they are not in the table.
ONLY_THE_FIRST = {"CHOOSECOLS", "CHOOSEROWS"}
AS_ARRAY = {"COVAR", "COVARIANCE.P", "HSTACK", "SUMPRODUCT", "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "VSTACK"}


def write_table(table, maxima):
    lines = [
        "//! The positions of a call where an error written in the formula is not an error to pass on but a refusal, as",
        "//! `LibreOffice` 26.8 answers (`OFFSET(#NAME?,0,0)` is `Err:504`).",
        "//!",
        "//! Generated by `sheetrs-compiler/scripts/probe_literal_error.py --write`; do not edit by hand.",
        "",
        "/// `(function, namespace, positions)`: a character per argument from the first, `.` where the function takes a",
        "/// value (an error written there is that error), and where it wants a reference or an array the answer to an",
        "/// error written there: `4` is `Err:504`, `2` is `Err:502`, `v` is `#VALUE!`, `7` is `Err:517`. A function with",
        "/// any number of arguments takes the rest as its last character says.",
        "#[rustfmt::skip]",
        "pub(super) const LITERAL_ERROR: &[(&str, &str, &str)] = &[",
    ]
    for (ns, name), positions in sorted(table.items(), key=lambda kv: (kv[0][1], kv[0][0])):
        if name in AS_ARRAY:
            continue
        text = "".join(positions.get(i, ".") for i in range(max(positions) + 1))
        if name in ONLY_THE_FIRST:
            text = text[0] + "."
        lines.append(f'    ("{name}", "{ns}", "{text}"),')
    lines.append("];")
    with open(OUT, "w", encoding="utf-8") as out:
        out.write("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()

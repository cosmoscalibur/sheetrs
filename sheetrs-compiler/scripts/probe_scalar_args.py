#!/usr/bin/env python3
"""Which arguments of which functions does LibreOffice refuse a plain value in?

    scripts/probe_scalar_args.py      # needs cargo, openpyxl and soffice ($SOFFICE: the 26.8 build)

For each function of the engine's registry (`probe_arity.find_spellings` says how an XLSX spells it) and
each position of a call with as many arguments as the function needs, a call with a range everywhere
(`Data!A1:A3`) is compared with the same call with a plain number (`5`) in that position. A position
where the range answers anything but an `Err:504`/`#REF!` and the number answers `Err:504`, `Err:502`
or `#REF!` is a candidate for `MATRIX_ARGS` of `sheetrs-engine/src/functions/mod.rs`, which is curated
from this list (an argument that takes a number as a number also answers `Err:502` to `5`, so the
candidates need a look: `CORREL(5, A1:A3)` is a difference of size, not a plain value).
"""
import collections
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import lo_probe  # noqa: E402
import probe_arity as pa  # noqa: E402

REFUSED = ("Err:504", "Err:502", "#REF!")


def main():
    funcs = [(ns, name, int(mn), int(mx)) for ns, name, mn, mx in pa.registry() if not name.startswith("_")]
    funcs = [f for f in funcs if (f[0], f[1]) not in pa.NOT_PROBED]
    spelling = pa.find_spellings(funcs)
    rows = []
    for ns, name, mn, mx in funcs:
        if (ns, name) not in spelling:
            continue
        prefix, nm = spelling[(ns, name)]
        n = max(mn, 1)
        base = [pa.R] * n
        rows.append((name, n, -1, f"{prefix}{nm}({','.join(base)})"))
        for i in range(n):
            args = list(base)
            args[i] = "5"
            rows.append((name, n, i, f"{prefix}{nm}({','.join(args)})"))
    answers = lo_probe.probe_all([r[3] for r in rows])
    seen = collections.defaultdict(dict)
    for r, a in zip(rows, answers):
        seen[(r[0], r[1])][r[2]] = a[1] if a[0] == "error" else "value"
    for (name, n), got in sorted(seen.items()):
        for i in range(n):
            if got[i] in REFUSED and got[-1] not in ("Err:504", "#REF!"):
                print(f"{name}\targument {i + 1} of {n}\t{got[i]}\t(ranges: {got[-1]})")


if __name__ == "__main__":
    main()

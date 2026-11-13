#!/usr/bin/env python3
"""Writes the function coverage report (docs/COVERAGE.md) to stdout.

Compares the engine's registry with the catalogue of the installed LibreOffice
(asked through UNO, so it is exactly what `soffice` can evaluate) and with the
formulas of the oracle fixtures in tests/assets.

Usage: scripts/function_coverage.py > docs/COVERAGE.md   (needs soffice, python3-uno,
openpyxl and a cargo toolchain)
"""
import collections
import json
import os
import re
import subprocess
import time
import zipfile

import openpyxl
import uno

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "..", "tests", "assets")
CATS = {1: "Database", 2: "Date and time", 3: "Financial", 4: "Information", 5: "Logical",
        6: "Mathematical", 7: "Array", 8: "Statistical", 9: "Lookup and spreadsheet", 10: "Text",
        11: "Add-in (engineering, analysis)"}


def libreoffice_functions():
    p = subprocess.Popen([os.environ.get("SOFFICE", "soffice"), "--headless", "--invisible", "--norestore",
                          "--accept=socket,host=localhost,port=2002;urp;"],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    ctx = None
    for _ in range(60):
        try:
            local = uno.getComponentContext()
            r = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
            ctx = r.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(1)
    smgr = ctx.ServiceManager
    fd = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionDescriptions", ctx)
    out = []
    for i in range(fd.Count):
        props = {x.Name: x.Value for x in fd.getByIndex(i)}
        out.append((props["Name"].upper(), CATS[props["Category"]]))
    version = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    try:
        version.terminate()
    except Exception:
        pass
    p.wait(timeout=30)
    return out


def our_functions():
    out = subprocess.run(["cargo", "run", "-q", "-p", "sheetrs-compiler", "--example", "list_functions"],
                         capture_output=True, text=True, check=True, cwd=HERE).stdout
    names = set()
    for line in out.splitlines():
        ns, name = line.split()
        names.add(name)
        # LibreOffice calls Excel's CEILING and FLOOR (the `microsoft` ones) CEILING.XCL / FLOOR.XCL.
        if ns == "microsoft" and name in ("CEILING", "FLOOR"):
            names.add(name + ".XCL")
    return names


# Resolved while compiling (`formula.rs`): `LET` by substitution, `ISFORMULA` and `FORMULA` of a
# literal reference from what the cell holds.
LOWERED = {"LET", "ISFORMULA", "FORMULA"}

# Listed by LibreOffice, but a formula that names them is `#NAME?` in an OpenDocument file, in an
# XLSX file (bare, `_xlfn.` and `_xlfn.ORG.LIBREOFFICE.`) and in LibreOffice's own grammar
# (probed with soffice 26.8): no file can use them.
UNREACHABLE = {
    "EFFECT_ADD", "NOMINAL_ADD", "CUMIPMT_ADD", "CUMPRINC_ADD", "ISEVEN_ADD", "ISODD_ADD",
    "GCD_EXCEL2003", "LCM_EXCEL2003", "WEEKNUM_EXCEL2003", "NETWORKDAYS_EXCEL2003", "CONVERT_OOO",
}


# The names Excel's file format gives to functions that LibreOffice (and an ODS) calls by another.
RENAMED_IN_FILES = {"FORECAST.ETS": "FORECAST.ETS.ADD", "FORECAST.ETS.CONFINT": "FORECAST.ETS.PI.ADD",
                    "FORECAST.ETS.STAT": "FORECAST.ETS.STAT.ADD"}


def oracle_formulas():
    forms = []
    wb = openpyxl.load_workbook(os.path.join(ASSETS, "semantics.xlsx"))
    forms += [c.value for ws in wb for r in ws.iter_rows() for c in r
              if isinstance(c.value, str) and c.value.startswith("=")]
    n_semantics = len(forms)
    for f in ["probe.ods", "probe_literal.ods", "probe_regex.ods"]:
        x = zipfile.ZipFile(os.path.join(ASSETS, f)).read("content.xml").decode()
        forms += re.findall(r'table:formula="([^"]*)"', x)
    return n_semantics, forms


def main():
    lo = libreoffice_functions()
    ours = our_functions() | LOWERED
    n_sem, forms = oracle_formulas()
    used = set()
    for f in forms:
        for n in re.findall(r"([A-Za-z_][A-Za-z0-9_.]*)\(", f):
            n = n.upper().replace("_XLFN.", "")
            for prefix in ("ORG.LIBREOFFICE.", "ORG.OPENOFFICE.", "COM.MICROSOFT.", "LEGACY."):
                n = n.removeprefix(prefix)
            used.add(n)
            used.add(RENAMED_IN_FILES.get(n, n))
    rows = collections.defaultdict(lambda: [0, 0, 0])
    missing = collections.defaultdict(list)
    untested = []
    unreachable = []
    for name, cat in lo:
        if name in UNREACHABLE:
            unreachable.append(name)
            continue
        r = rows[cat]
        r[0] += 1
        if name in ours:
            r[1] += 1
            if name in used:
                r[2] += 1
            else:
                untested.append(name)
        else:
            missing[cat].append(name)
    tot = [sum(r[i] for r in rows.values()) for i in range(3)]
    print("# Function and operator coverage\n")
    print("Generated by `scripts/function_coverage.py` against "
          f"{subprocess.run([os.environ.get('SOFFICE', 'soffice'), '--version'], capture_output=True, text=True).stdout.strip()}.\n")
    print("The reference catalogue is the one LibreOffice reports through its own API "
          "(`FunctionDescriptions`): every function `soffice` can evaluate, which includes the ones "
          "of OpenDocument (OpenFormula) and the Excel-compatible ones. *Implemented* means in "
          "the engine's registry; *checked* means that an oracle formula (a formula of "
          "`tests/assets/semantics.*` or the probes, whose values LibreOffice computed) "
          "uses it.\n")
    print(f"**{tot[1]} of {tot[0]} LibreOffice functions are implemented (not counting the {len(unreachable)} that no file can use); {tot[2]} of them appear in "
          f"an oracle formula.** The oracle has {n_sem} formulas in `semantics.*` plus the probes; "
          "most functions have one or two formulas, not an exhaustive set of edge cases.\n")
    print("| Category | LibreOffice | Implemented | Checked against soffice |")
    print("| --- | ---: | ---: | ---: |")
    for cat, r in sorted(rows.items(), key=lambda kv: -kv[1][0]):
        print(f"| {cat} | {r[0]} | {r[1]} | {r[2]} |")
    print(f"| **Total** | **{tot[0]}** | **{tot[1]}** | **{tot[2]}** |\n")
    print("## Not implemented, by category\n")
    print("A workbook that calls any of them does not build: the error names the function and the first cell. This list is the inventory the build checks against (`CATALOGUE` in `functions/namespaces.rs`).\n")
    for cat, names in sorted(missing.items(), key=lambda kv: -len(kv[1])):
        print(f"* **{cat}** ({len(names)}): " + ", ".join(f"`{n}`" for n in names))
    print("\n## Listed by LibreOffice but unusable in a file\n")
    print("A formula that names them is `#NAME?` in an OpenDocument file, in an XLSX file (bare, with `_xlfn.` or with `_xlfn.ORG.LIBREOFFICE.`) and in LibreOffice's own grammar, so no workbook can call them and they are not counted above: " + ", ".join(f"`{n}`" for n in sorted(unreachable)) + ".")
    print("\n## Resolved while compiling\n")
    print("`LET` (by substitution), and `ISFORMULA` and `FORMULA` of a literal reference (from what the cell holds) are resolved by the compiler, not by the engine's registry; `ISFORMULA` and `FORMULA` of a reference computed at run time stop the build.")
    print("\n## Implemented but not checked against soffice\n")
    print(", ".join(f"`{n}`" for n in sorted(untested)) or "none")
    print("\nFunctions are registered by their LibreOffice name; the names a file gives them are in "
          "`docs/FUNCTION_NAMES.md` (`ISNUMBER` counts booleans as numbers, and "
          "`EASTERSUNDAY`).\n")
    print("## Operators and syntax\n")
    print("Implemented: `+ - * / ^ &`, the comparisons `= <> < <= > >=`, unary `-` and `+`, "
          "postfix `%`, the range operator `:` (also between computed references, as in "
          "`A1:INDEX(...)`), whole-column and whole-row ranges, sheet-qualified references, "
          "defined names, array constants and `{...}` array formulas. In the oracle: every "
          "arithmetic operator, `&`, `%` and `:`, `=`, `<>`, `<`, `>` and `>=` (some with a "
          "single formula; `<=` appears only in probes).\n")
    print("Not implemented (a build stops, naming the formula): reference **intersection** (space) and **union** "
          "(`,` inside parentheses in Excel, `~` in OpenDocument), the **spill** operator "
          "`A1#`, **structured references** to tables (`Table1[Col]`), and **external workbook** "
          "references (`[Book.xlsx]Sheet!A1`).\n")
    excel_new = ("XMATCH FILTER SORT SORTBY UNIQUE SEQUENCE RANDARRAY LET LAMBDA MAP REDUCE SCAN BYROW "
                 "BYCOL MAKEARRAY TEXTBEFORE TEXTAFTER TEXTSPLIT VSTACK HSTACK TAKE DROP CHOOSECOLS "
                 "CHOOSEROWS TOCOL TOROW WRAPCOLS WRAPROWS EXPAND ARRAYTOTEXT VALUETOTEXT GROUPBY PIVOTBY").split()
    known = {n for n, _ in lo}
    absent = [n for n in excel_new if n not in known]
    print("## Excel functions this LibreOffice does not have\n")
    print("Checked against the catalogue above (the list of Excel functions is from the Excel "
          "documentation, not generated): " + (", ".join(f"`{n}`" for n in absent) or "none") + ".\n")
    print("## Behavior parity\n")
    print("Behavior is checked only where `soffice` can answer. The reference is LibreOffice for every file "
          "(also XLSX); the oracle is its own results. Where it differs from Excel, the engine gives LibreOffice's answer, "
          "documented in [FIDELITY.md](FIDELITY.md) next to Excel's documented behavior, which is not checked against "
          "Excel itself: no Excel is available here.")


if __name__ == "__main__":
    main()

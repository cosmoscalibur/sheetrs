#!/usr/bin/env python3
"""Writes the engine's locale table (sheetrs-engine/src/locale_data.rs) from LibreOffice.

For each locale of LOCALES, a LibreOffice running with that regional setting is asked for
  * its locale data (UNO `LocaleData2`): separators, month and day names (full, abbreviated,
    and the abbreviations), the patterns it accepts for a date typed as text, the date and currency formats;
  * how its input scanner reads a date with a month name: 31 probes (`DATEVALUE`) over the forms
    `Mon D Y`, `D Mon Y`, `Y Mon D`, `Mon D`, `D Mon`, `Mon Y`, `Y Mon`, with two-digit years, with
    each separator, a dot after the day, and a comma; the result of each probe
    is one letter of the `named` string (see `NAMED_FORMS`);
  * the texts it shows for TRUE, FALSE, AM and PM.
Nothing is derived from ICU or from a language tag: each locale's data is what LibreOffice says.

Run with LibreOffice's Python (a few seconds per locale). With the locale tags as arguments it
probes only those, but the file then holds only them:
  SOFFICE=soffice26 PYTHONPATH=/usr/lib/python3/dist-packages \
    .../program/python scripts/gen_locale_table.py ../sheetrs-engine/src/locale_data.rs [TAG ...]
"""
import datetime
import os
import re
import subprocess
import sys
import tempfile
import time
import unicodedata

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import Locale

SOFFICE = os.environ.get("SOFFICE", "soffice")
OUT = sys.argv[1] if len(sys.argv) > 1 else "locale_data.rs"

# The locales of the table: those the engine is checked against LibreOffice for, row by row
# (a fixture `semantics_<tag>` in tests/assets, see scripts/gen_semantics_fixture.py). A locale
# is added here only together with its fixture and once the engine reproduces it; the data of a
# locale is not support for it. The first of each language is what the bare language tag means.
LOCALES = ["en-US", "en-GB", "es-ES", "es-CO", "es-MX"]

# Day 13 and year 2024 cannot be mistaken for a month, so the result of a probe says which role
# LibreOffice gave each number: `M` the month (March, written as the locale does), `A` 13, `B` 2024,
# `b` 25. For each separator of SEPS the forms of FORMS come,
# then the dot-after-day and comma forms of EXTRA (written with their own separators).
SEPS = [" ", "-", "/"]
FORMS = ["M A B", "A M B", "B M A", "M A", "A M", "M B", "B M", "M A b", "A M b"]
EXTRA = ["A. M B", "A. M", "M A, B"]
NAMED_FORMS = [f"{f} [{s}]" for s in SEPS for f in FORMS] + EXTRA


def pv(n, v):
    p = PropertyValue()
    p.Name, p.Value = n, v
    return p


def start(locale, port):
    prof = tempfile.mkdtemp(prefix="lo_loc_")
    os.makedirs(prof + "/user")
    open(prof + "/user/registrymodifications.xcu", "w").write(
        '<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry" '
        'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse">'
        f"<value>{locale}</value></prop></item></oor:items>"
    )
    proc = subprocess.Popen(
        [SOFFICE, "--headless", "--invisible", "--norestore", f"-env:UserInstallation=file://{prof}",
         f"--accept=socket,host=localhost,port={port};urp;"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    local = uno.getComponentContext()
    for _ in range(120):
        try:
            r = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
            ctx = r.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
            return proc, ctx
        except Exception:
            time.sleep(1)
    raise SystemExit("LibreOffice did not start")


def evaluate(doc, formulas):
    """Result of each formula: ('n', number) or ('s', text) or ('e', error code)."""
    sh = doc.Sheets.getByIndex(0)
    for i, f in enumerate(formulas):
        sh.getCellByPosition(0, i).setFormula(f)
    doc.calculateAll()
    out = []
    for i in range(len(formulas)):
        c = sh.getCellByPosition(0, i)
        if c.getError():
            out.append(("e", c.getError()))
        elif c.FormulaResultType2 == 2:
            out.append(("s", c.getString()))
        else:
            out.append(("n", c.getValue()))
    return out


def serial(d):
    return (d - datetime.date(1899, 12, 30)).days


def named_class(form, result, this_year):
    """One letter: `-` rejected, `d` the number `A` is the day, `y` it is the year, `?` anything else.

    With a year of two digits (`b`) the day may be the last number instead (`y`: A is the year).
    A form without `B` or `b` has no year of its own: LibreOffice gives the day of this year, or
    reads `A` as the year of the first of the month."""
    kind, value = result
    if kind != "n":
        return "-"
    tokens = form.replace(",", "").replace(".", "").split()
    if "A" not in tokens:
        return "d" if value == serial(datetime.date(2024, 3, 1)) else "?"
    if "B" not in tokens and "b" not in tokens:
        if value == serial(datetime.date(this_year, 3, 13)):
            return "d"
        return "y" if value == serial(datetime.date(2013, 3, 1)) else "?"
    if "b" in tokens:
        if value == serial(datetime.date(2025, 3, 13)):
            return "d"
        return "y" if value == serial(datetime.date(2013, 3, 25)) else "?"
    return "d" if value == serial(datetime.date(2024, 3, 13)) else "?"


def rust_str(s):
    out = []
    for ch in s:
        if ch in '"\\':
            out.append("\\" + ch)
        elif unicodedata.category(ch)[0] in "ZC" and ch != " " or ord(ch) < 32:
            out.append("\\u{%x}" % ord(ch))
        else:
            out.append(ch)
    return '"' + "".join(out) + '"'


def rust_char(ch):
    return "'" + rust_str(ch)[1:-1].replace("'", "\\'") + "'"


# Microsoft's language ids, which `[$-409]` in a format code names. LibreOffice reads the ones it has
# a locale for (`C0A` is Spanish of Spain; its traditional-sort `40A` it does not know, and TEXT then
# keeps the locale it runs with). Each is checked below against a LibreOffice running in en-US.
LCIDS = {"en-US": 0x409, "en-GB": 0x809, "es-ES": 0xC0A, "es-CO": 0x240A, "es-MX": 0x80A}


def honored(en, lcid, days, months):
    """Whether TEXT, in an en-US LibreOffice, writes the names of the locale `lcid` names."""
    got = evaluate(en, [f'=TEXT(45367.75;"[$-{lcid:X}]dddd mmmm")'])[0]
    return got == ("s", f"{days[6]} {months[2]}")


def date_order(doc, loc):
    """The order of day, month and year in the locale's short date (3/21/2013 -> mdy)."""
    formats = doc.NumberFormats
    key = formats.getFormatIndex(uno.getConstantByName("com.sun.star.i18n.NumberFormatIndex.DATE_SYSTEM_SHORT"), loc)
    cell = doc.Sheets.getByIndex(0).getCellByPosition(5, 0)
    cell.setValue(serial(datetime.date(2013, 3, 21)))
    cell.NumberFormat = key
    text = cell.getString()
    where = {"d": text.find("21"), "m": text.replace("2013", "####").find("03"), "y": text.find("13")}
    where["m"] = where["m"] if where["m"] >= 0 else text.find("3")
    where["y"] = text.find("2013") if "2013" in text else where["y"]
    return "".join(sorted(where, key=where.get))


def currency(pos, neg, item):
    """The currency symbol and the templates (`¤` symbol, `#` number) of DOLLAR's two outputs."""
    digits = lambda t: re.search(r"\d[\d.,\u00a0\u202f' ]*\d|\d", t)

    def tpl(t):
        m = digits(t)
        return t[: m.start()] + "#" + t[m.end():]

    pos_t, neg_t = tpl(pos), tpl(neg)
    symbol = (pos_t.replace("#", "").strip().strip("\u00a0\u202f"))
    return symbol, pos_t.replace(symbol, "¤"), neg_t.replace(symbol, "¤")


def probe(tag, en):
    port = 2100 + LOCALES.index(tag)
    proc, ctx = start(tag, port)
    try:
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))
        ld = smgr.createInstanceWithContext("com.sun.star.i18n.LocaleData2", ctx)
        lang, region = tag.split("-")
        loc = Locale(lang, region, "")
        item = ld.getLocaleItem2(loc)
        cal = ld.getAllCalendars2(loc)[0]
        months = [m.FullName for m in cal.Months]
        days = [d.FullName for d in cal.Days]
        lcid = LCIDS[tag]
        if not tag.startswith("en-") and not honored(en, lcid, days, months):
            print(f"{tag}: LibreOffice does not read [$-{lcid:X}]", file=sys.stderr)
            lcid = 0
        words = ld.getReservedWord(loc)
        # AM and PM as LibreOffice writes them; DOLLAR shows its currency format.
        texts = evaluate(doc, ['=TEXT(0.25;"AM/PM")', '=TEXT(0.75;"AM/PM")', "=DOLLAR(1234.5;2)", "=DOLLAR(-1234.5;2)"])
        order = date_order(doc, loc)
        symbol, pos, neg = currency(texts[2][1], texts[3][1], item)
        this_year = datetime.date.today().year
        mon = months[2]
        subs = {"M": mon, "A": "13", "B": "2024", "b": "25"}
        probes = [(f, s.join(subs[t] for t in f.split())) for s in SEPS for f in FORMS]
        for f in EXTRA:
            text = " ".join(subs[t.rstrip(".,")] + t[len(t.rstrip(".,")):] for t in f.split())
            probes.append((f, text))
        results = evaluate(doc, [f'=DATEVALUE("{t}")' for _, t in probes])
        named = "".join(named_class(f, r, this_year) for (f, _), r in zip(probes, results))
        row = {
            "tag": tag, "lcid": lcid, "order": order,
            "decimal": item.decimalSeparator, "group": item.thousandSeparator, "date_sep": item.dateSeparator,
            "true": words[0].upper(), "false": words[1].upper(), "am": texts[0][1], "pm": texts[1][1],
            "symbol": symbol, "pos": pos, "neg": neg,
            "months": months, "months_abbr": [m.AbbrevName for m in cal.Months],
            "days": [d.FullName for d in cal.Days], "days_abbr": [d.AbbrevName for d in cal.Days],
            "patterns": list(ld.getDateAcceptancePatterns(loc)),
            "named": named,
        }
        doc.close(True)
        return row
    finally:
        proc.terminate()
        try:
            proc.wait(15)
        except Exception:
            proc.kill()


def emit(rows):
    arr = lambda xs: "[" + ", ".join(rust_str(x) for x in xs) + "]"
    out = [
        "//! Regional settings of the locales LibreOffice offers, as LibreOffice reports them.",
        "//! Generated by `sheetrs-compiler/scripts/gen_locale_table.py` (LibreOffice "
        + os.environ.get("LO_VERSION", "26.8") + "); do not edit.",
        "",
        "use crate::options::LocaleRow;",
        "",
        "/// The locale a bare language tag means.",
        "pub const DEFAULT_REGION: &[(&str, &str)] = &[",
    ]
    seen = []
    for r in rows:
        lang = r["tag"].split("-")[0]
        if lang not in seen:
            seen.append(lang)
            out.append(f'    ("{lang}", "{r["tag"]}"),')
    out += ["];", "", f"/// Forms of a date with a month name, in the order of the `named` string: {', '.join(NAMED_FORMS)}.",
            "pub const LOCALES: &[LocaleRow] = &["]
    for r in rows:
        out += [
            "    LocaleRow {",
            f'        tag: "{r["tag"]}",',
            f'        lcid: 0x{r["lcid"]:X},',
            f'        order: "{r["order"]}",',
            f'        decimal: {rust_char(r["decimal"])},',
            f'        group: {rust_char(r["group"])},',
            f'        date_sep: {rust_char(r["date_sep"])},',
            f'        true_text: {rust_str(r["true"])},',
            f'        false_text: {rust_str(r["false"])},',
            f'        am: {rust_str(r["am"])},',
            f'        pm: {rust_str(r["pm"])},',
            f'        currency: {rust_str(r["symbol"])},',
            f'        currency_pos: {rust_str(r["pos"])},',
            f'        currency_neg: {rust_str(r["neg"])},',
            f'        months: {arr(r["months"])},',
            f'        months_abbr: {arr(r["months_abbr"])},',
            f'        days: {arr(r["days"])},',
            f'        days_abbr: {arr(r["days_abbr"])},',
            f'        date_patterns: &{arr(r["patterns"])},',
            f'        named: "{r["named"]}",',
            "    },",
        ]
    out.append("];")
    open(OUT, "w").write("\n".join(out) + "\n")


def main():
    tags = [t for t in sys.argv[2:]] or LOCALES
    rows = []
    proc, ctx = start("en-US", 2099)
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    en = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))
    for t in tags:
        if t not in LOCALES:
            LOCALES.append(t)
        row = probe(t, en)
        print(t, row["named"], row["order"], row["patterns"], file=sys.stderr)
        rows.append(row)
    proc.terminate()
    emit(rows)


if __name__ == "__main__":
    main()

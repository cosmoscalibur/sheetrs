#!/usr/bin/env python3
"""Makes the workbooks of the test fixtures byte-for-byte reproducible.

The generators of the fixtures (`gen_*.py`, `make_*.py`) build workbooks with openpyxl and have
LibreOffice recalculate and save them. Every run writes the time of the run into the zip entries and
the document properties, so that running a generator twice changes bytes without changing a single
cell, and the committed binary files show up as modified. `normalize` rewrites a workbook so that what
depends on the moment of the run is a constant:

  * the date of every zip entry (`FIXED_ZIP_DATE`), keeping the order, the compression method and,
    for an ODS, `mimetype` first and stored;
  * `docProps/core.xml` (`dcterms:created`, `dcterms:modified`) of an XLSX, except the modification date of the fixtures that
    are about it (`DELIBERATE_MODIFIED`);
  * `meta.xml` (`meta:creation-date`, `dc:date`, `meta:editing-duration`) of an ODS;
  * the date and the time LibreOffice writes for a date field of a header or footer
    (`text:date-value`, `text:time-value`) in `styles.xml` and `content.xml` of an ODS, and the date of
    each note (`dc:date` in `content.xml`);
  * the order of the master pages of an ODS (`styles.xml`), which LibreOffice takes from a hash table;
  * the ids of the axes of the charts of an XLSX, which LibreOffice draws at random;
  * with `freeze_volatile=True`, what a formula that is just `NOW()`, `TODAY()` or `RAND()` has cached
    (2024-01-01 00:00:00, 0.5): LibreOffice recalculates it when it opens the file. Off by default: the
    cached `TODAY()` of the semantics fixtures is the clock their oracle gives the engine.

Cell content, formulas, cached values and styles are not touched. What stays is the version of the
application that wrote the file (`Application` in `docProps/app.xml`, `meta:generator`): the output is
the same for the same LibreOffice build, and a new build is a real change that shows in the diff.

Every generator calls `normalize(path)` on each workbook it leaves in the repository. As a command:

    scripts/fixture_determinism.py FILE...        normalize the files in place
    scripts/fixture_determinism.py --check        list the tracked fixtures that are not normalized
                                                  (exit code 1 if there is any); with FILE... only those

The command needs only the standard library.
"""
import os
import re
import stat
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

FIXED_ZIP_DATE = (2024, 1, 1, 0, 0, 0)
FIXED_ISO = "2024-01-01T00:00:00"
FIXED_DATE = "2024-01-01"
FIXED_TIME = "00:00:00"
COMPRESS_LEVEL = 6

# The fixtures whose modification date is part of what they test (a workbook that was not touched in
# years): path relative to the repository -> the date, `YYYY-MM-DDTHH:MM:SS`. `created` is still fixed.
DELIBERATE_MODIFIED = {
    "tests/lint_book.xlsx": "2020-01-15T10:00:00",
    "tests/lint_book.ods": "2020-01-15T10:00:00",
}
_ROOT = Path(__file__).resolve().parent.parent


def _modified_of(path):
    try:
        rel = Path(path).resolve().relative_to(_ROOT).as_posix()
    except ValueError:
        return FIXED_ISO
    return DELIBERATE_MODIFIED.get(rel, FIXED_ISO)


def _patches(modified):
    """(entry name -> [(pattern, replacement)]) applied to the bytes of the entry."""
    fixed, mod = FIXED_ISO.encode(), modified.encode()
    xlsx = [
        (re.compile(rb"(<dcterms:created\b[^>]*>)[^<]*(</dcterms:created>)"), rb"\g<1>" + fixed + rb"Z\g<2>"),
        (re.compile(rb"(<dcterms:modified\b[^>]*>)[^<]*(</dcterms:modified>)"), rb"\g<1>" + mod + rb"Z\g<2>"),
    ]
    ods = [
        (re.compile(rb"(<meta:creation-date>)[^<]*(</meta:creation-date>)"), rb"\g<1>" + fixed + rb"\g<2>"),
        (re.compile(rb"(<dc:date>)[^<]*(</dc:date>)"), rb"\g<1>" + mod + (b".000000000" if modified != FIXED_ISO else b"") + rb"\g<2>"),
        (re.compile(rb"(<meta:editing-duration>)[^<]*(</meta:editing-duration>)"), rb"\g<1>PT0S\g<2>"),
    ]
    fields = [
        (re.compile(rb'text:date-value="[^"]*"'), b'text:date-value="' + FIXED_DATE.encode() + b'"'),
        (re.compile(rb'text:time-value="[^"]*"'), b'text:time-value="' + FIXED_TIME.encode() + b'"'),
    ]
    # the date of a note (`office:annotation`) is the moment the generator added it
    notes = [(re.compile(rb"(<dc:date>)[^<]*(</dc:date>)"), rb"\g<1>" + fixed + rb"\g<2>")]
    return {"docProps/core.xml": xlsx, "meta.xml": ods, "styles.xml": fields, "content.xml": fields + notes}


_WORKBOOK_SUFFIXES = (".xlsx", ".xlsm", ".ods")

# What a formula with a volatile function caches is the moment (or the random number) of the run:
# 2024-01-01 00:00:00 for NOW and TODAY, 0.5 for RAND. LibreOffice recalculates them when it opens the
# file, so no test can depend on the cached value.
_SERIAL_FIXED = "45292"  # 2024-01-01 as a serial number
# Only a formula that is nothing but a volatile call (or `IF(1,NOW(),RAND())`, which gives NOW): a formula
# such as `RANDBETWEEN(2,2)` or `RAND()*0` caches a result that does not depend on the run, and the
# oracle compares the engine with it.
_VOLATILE_SHAPE = re.compile(rb"^(NOW\(\)|TODAY\(\)|RAND\(\)|IF\(1,(NOW|TODAY)\(\),RAND\(\)\))$")


def _volatile_kind(formula):
    """b'now', b'today', b'rand' or None for the text of a formula of the file."""
    f = formula.upper().replace(b" ", b"").replace(b";", b",").replace(b"&apos;", b"")
    for prefix in (b"OF:=", b"="):
        if f.startswith(prefix):
            f = f[len(prefix):]
            break
    f = f.replace(b"_XLFN.", b"")
    if not _VOLATILE_SHAPE.match(f):
        return None
    return b"now" if b"NOW(" in f else b"today" if b"TODAY(" in f else b"rand"
_XLSX_VOLATILE = re.compile(rb"(<c\b[^>]*>\s*<f\b[^>]*>)([^<]*)(</f>\s*<v>)([^<]*)(</v>)")
_ODS_VOLATILE = re.compile(
    rb'(<table:table-cell\b[^>]*?table:formula="([^"]*)"[^>]*>)<text:p>([^<]*)</text:p>', re.S)
_DISPLAY_SHAPES = [  # (shape of the text a cell shows, what it shows for 2024-01-01 00:00:00)
    (re.compile(rb"^\d\d/\d\d/\d\d \d\d:\d\d [AP]M$"), b"01/01/24 12:00 AM"),
    (re.compile(rb"^\d\d/\d\d/\d\d$"), b"01/01/24"),
    (re.compile(rb"^\d{4}-\d\d-\d\d$"), b"2024-01-01"),
    (re.compile(rb"^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d$"), b"2024-01-01T00:00:00"),
]


def _freeze_xlsx(data):
    def freeze(m):
        kind = _volatile_kind(m.group(2))
        if kind in (b"now", b"today"):
            value = _SERIAL_FIXED.encode()
        elif kind == b"rand":
            value = b"0.5"
        else:
            return m.group(0)
        return m.group(1) + m.group(2) + m.group(3) + value + m.group(5)

    return _XLSX_VOLATILE.sub(freeze, data)


def _freeze_ods(data):
    def freeze(m):
        tag, kind, text = m.group(1), _volatile_kind(m.group(2)), m.group(3)
        if kind in (b"now", b"today"):
            fixed = b"2024-01-01" if kind == b"today" else FIXED_ISO.encode()
            tag = re.sub(rb'office:date-value="[^"]*"', b'office:date-value="' + fixed + b'"', tag)
            for shape, shown in _DISPLAY_SHAPES:
                if shape.match(text):
                    text = shown
                    break
            else:
                raise ValueError(f"fixture_determinism: the display of a volatile cell has a shape not known: {text!r}")
        elif kind == b"rand":
            tag = re.sub(rb'office:value="[^"]*"', b'office:value="0.5"', tag)
            text = b"0.5"
        else:
            return m.group(0)
        return tag + b"<text:p>" + text + b"</text:p>"

    return _ODS_VOLATILE.sub(freeze, data)

# LibreOffice gives the axes of a chart of an XLSX random ids (`c:axId`, `c:crossAx`); they are
# renumbered in the order they appear, which keeps the references between axes.
_CHART_ENTRY = re.compile(r"^xl/charts/chart\d+\.xml$")
_AXIS_ID = re.compile(rb'(<c:(?:axId|crossAx) val=")(\d+)(")')


def _renumber_axes(data):
    seen = {}

    def number(m):
        return m.group(1) + str(seen.setdefault(m.group(2), 500000000 + len(seen))).encode() + m.group(3)

    return _AXIS_ID.sub(number, data)

# LibreOffice writes the master pages of an ODS (`PageStyle_Calc`, `PageStyle_Wide`...) in the order of
# an internal hash table, which changes from run to run. Nothing refers to their position, only to
# their name, so they are written sorted by name.
_MASTER_STYLES = re.compile(rb"(<office:master-styles>)(.*?)(</office:master-styles>)", re.S)
_MASTER_PAGE = re.compile(rb"<style:master-page\b.*?</style:master-page>|<style:master-page\b[^>]*/>", re.S)
_MASTER_NAME = re.compile(rb'style:name="([^"]*)"')


def _sort_master_pages(data):
    def sort(m):
        pages = _MASTER_PAGE.findall(m.group(2))
        if len(pages) < 2 or b"".join(pages) != m.group(2):
            return m.group(0)  # something else between them: leave it as it is
        pages.sort(key=lambda p: _MASTER_NAME.search(p).group(1))
        return m.group(1) + b"".join(pages) + m.group(3)

    return _MASTER_STYLES.sub(sort, data)


def _normalized_entries(path, freeze_volatile=False):
    """(info, data) of every entry of the zip, patched, in their order; raises BadZipFile."""
    patches = _patches(_modified_of(path))
    with zipfile.ZipFile(path) as zin:
        for info in zin.infolist():
            data = zin.read(info)
            patched = data
            for pattern, replacement in patches.get(info.filename, ()):
                patched = pattern.sub(replacement, patched)
            if patched != data:
                ET.fromstring(patched)  # a pattern that breaks the XML must fail here, not in a reader
                data = patched
            if info.filename == "styles.xml":
                data = _sort_master_pages(data)
            if _CHART_ENTRY.match(info.filename):
                data = _renumber_axes(data)
            if freeze_volatile and re.match(r"^xl/worksheets/sheet\d+\.xml$", info.filename):
                data = _freeze_xlsx(data)
            if freeze_volatile and info.filename == "content.xml":
                data = _freeze_ods(data)
            yield info, data


def normalize(path, freeze_volatile=False):
    """Rewrites the workbook `path` in place; returns True when its bytes changed.

    `freeze_volatile` also sets what the cells that are just `NOW()`, `TODAY()` or `RAND()` cache to a
    constant. It is for the fixtures whose volatile cells have no other cell that depends on them: the
    semantics fixtures keep the `TODAY()` of their run, because the oracle takes the clock of the engine
    from it.

    A file that is not a zip (a fixture that is corrupt on purpose) is left as it is.
    """
    path = Path(path)
    try:
        entries = list(_normalized_entries(path, freeze_volatile))
    except zipfile.BadZipFile:
        return False
    original = path.read_bytes()
    with tempfile.NamedTemporaryFile(dir=path.parent, suffix=".tmp", delete=False) as tmp:
        tmp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(tmp_path, "w") as zout:
            for info, data in entries:
                out = zipfile.ZipInfo(info.filename, date_time=FIXED_ZIP_DATE)
                out.compress_type = info.compress_type
                out.create_system = info.create_system
                out.external_attr = info.external_attr
                zout.writestr(out, data, compress_type=info.compress_type, compresslevel=COMPRESS_LEVEL)
        changed = tmp_path.read_bytes() != original
        if changed:
            os.chmod(tmp_path, stat.S_IMODE(path.stat().st_mode))  # the temporary file is 0600
            tmp_path.replace(path)
        else:
            tmp_path.unlink()
    except BaseException:
        tmp_path.unlink(missing_ok=True)
        raise
    return changed


def problems(path):
    """What of `path` still depends on the moment it was generated; [] when it is normalized."""
    path = Path(path)
    try:
        entries = list(_normalized_entries(path))
        with zipfile.ZipFile(path) as zin:
            infos = zin.infolist()
            stored = {i.filename: zin.read(i) for i in infos}
    except zipfile.BadZipFile:
        return []
    found = []
    if any(i.date_time != FIXED_ZIP_DATE for i in infos):
        found.append("zip entry dates")
    for info, data in entries:
        if stored[info.filename] != data:
            found.append(info.filename)
    return found


def tracked_fixtures(root):
    out = subprocess.run(["git", "-C", str(root), "ls-files"], check=True, capture_output=True, text=True).stdout
    return [root / f for f in out.splitlines() if f.endswith(_WORKBOOK_SUFFIXES)]


def _main(argv):
    check = "--check" in argv
    files = [Path(a) for a in argv if not a.startswith("--")]
    root = Path(__file__).resolve().parent.parent
    if check:
        bad = 0
        for f in files or tracked_fixtures(root):
            found = problems(f)
            if found:
                bad += 1
                print(f"{f}: not normalized ({', '.join(found)})")
        if bad:
            print(f"{bad} fixture(s) depend on the time they were generated: scripts/fixture_determinism.py FILE...")
        return 1 if bad else 0
    if not files:
        print(__doc__)
        return 2
    for f in files:
        print(("normalized " if normalize(f) else "already normalized ") + str(f))
    return 0


if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))

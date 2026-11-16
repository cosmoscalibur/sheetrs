"""Which formulas of an ODS LibreOffice 26.8 calculates when it opens the file, instead of showing the result the file caches: the evidence
of `sheetrs_io::reader::ods_parser::recalculated_on_load` (and of the roots `recalculated_on_load` of `sheetrs_compiler::Cache`).

Each file is written by hand (no application made it), with a stale result in each formula cell (999, 9.99, false, the text S, a date of 1999)
that is not the one the formula gives. LibreOffice opens it and a cell that shows the result of the formula was calculated on load; one that
shows the stale one kept the cache. Nothing is saved: this is the default load.

    python3 probe_ods_recalculation.py gen DIR
    /opt/libreoffice26.8/program/python probe_ods_recalculation.py run DIR     # LibreOffice listening on port 2002

The tables (`R` calculated on load, `.` kept) are what `docs/approved-differences.md` and `FIDELITY.md` (*What LibreOffice calculates on load*) say:

* `generators`: one bare cell of every type, in files whose `meta:generator` is LibreOffice 26.8, 7.6, 4.2, OpenOffice, Excel, another tool or absent, and
  `office:version` 1.3, 1.2 or absent: the same row for all. The generator and the version do not decide it.
* `placement`: where the style of the cell comes from (the cell, the column, the row, a parent) and the type of the result.
* `numbers`: which `number:number-style` LibreOffice takes for General.
* `extra`: no type, array formulas, repeated cells.
"""
import glob
import itertools
import os
import sys
import zipfile

NS = ('xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" '
      'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
      'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
      'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" '
      'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" '
      'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
      'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
      'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ')
GENERATORS = {
    "lo26": "LibreOffice/26.8.0.0$Linux_X86_64 LibreOffice_project/abc",
    "lo7": "LibreOffice/7.6.4.1$Linux_X86_64 LibreOffice_project/60$Build-1",
    "lo4": "LibreOffice/4.2.8.2$Linux_X86_64 LibreOffice_project/420m0$Build-2",
    "ooo": "OpenOffice/4.1.14$Linux_X86_64 OpenOffice.org_project/414m1$Build-9808",
    "excel": "Microsoft Excel",
    "foreign": "SomeTool/1.0",
    "none": None,
}
CX = ' calcext:value-type="{}"'
# (label, formula, the stale cached result, what the formula gives: a number or a text)
TYPES = [
    ("float", "of:=1+1", 'office:value-type="float" office:value="999"' + CX.format("float"), 2),
    ("percent", "of:=1+1", 'office:value-type="percentage" office:value="9.99"' + CX.format("percentage"), 2),
    ("currency", "of:=1+1", 'office:value-type="currency" office:currency="USD" office:value="999"' + CX.format("currency"), 2),
    ("boolean", "of:=1=1", 'office:value-type="boolean" office:boolean-value="false"' + CX.format("boolean"), 1),
    ("string", "of:=CONCATENATE(&quot;a&quot;;&quot;b&quot;)", 'office:value-type="string" office:string-value="S"' + CX.format("string"), "ab"),
    ("date", "of:=DATE(2020;1;1)", 'office:value-type="date" office:date-value="1999-01-01"' + CX.format("date"), 43831),
    ("time", "of:=TIME(1;0;0)", 'office:value-type="time" office:time-value="PT05H00M00S"' + CX.format("time"), 1 / 24),
    ("error", "of:=1+1", 'office:value-type="string" office:string-value="" calcext:value-type="error"', 2),
    ("float,text", "of:=CONCATENATE(&quot;a&quot;;&quot;b&quot;)", 'office:value-type="float" office:value="999"' + CX.format("float"), "ab"),
    ("string,num", "of:=1+1", 'office:value-type="string" office:string-value="S"' + CX.format("string"), 2),
    ("no type", "of:=1+1", "", 2),
]
NUMBERS = {  # the number styles, by the `number:number` they hold (and what is around it)
    "n0": '<number:number number:min-integer-digits="1"/>',
    "min3": '<number:number number:min-integer-digits="3"/>',
    "min0": '<number:number number:min-integer-digits="0"/>',
    "group": '<number:number number:grouping="true" number:min-integer-digits="1"/>',
    "factor": '<number:number number:min-integer-digits="1" number:display-factor="1000"/>',
    "bare": "<number:number/>",
    "dec0": '<number:number number:decimal-places="0" number:min-integer-digits="1"/>',
    "dec2": '<number:number number:decimal-places="2" number:min-integer-digits="1"/>',
    "dec2only": '<number:number number:decimal-places="2"/>',
    "unit": '<number:number number:min-integer-digits="1"/><number:text> kg</number:text>',
    "minus": '<number:text>-</number:text><number:number number:min-integer-digits="1"/>',
}
OTHER = {
    "pct": '<number:percentage-style style:name="N_pct"><number:number number:decimal-places="0" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style>',
    "cur": '<number:currency-style style:name="N_cur"><number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol><number:number number:decimal-places="2" number:min-integer-digits="1"/></number:currency-style>',
    "date": '<number:date-style style:name="N_date"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>',
    "time": '<number:time-style style:name="N_time"><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/></number:time-style>',
    "bool": '<number:boolean-style style:name="N_bool"><number:boolean/></number:boolean-style>',
    "text": '<number:text-style style:name="N_text"><number:text-content/></number:text-style>',
}
AUTO = "".join(f'<number:number-style style:name="N_{k}">{v}</number:number-style>' for k, v in NUMBERS.items()) + "".join(OTHER.values())
AUTO += '<style:style style:name="s_empty" style:family="table-cell"/>'
AUTO += "".join(f'<style:style style:name="s_{k}" style:family="table-cell" style:data-style-name="N_{k}"/>' for k in list(NUMBERS) + list(OTHER))
AUTO += '<style:style style:name="s_par" style:family="table-cell" style:parent-style-name="Fmt"/><style:style style:name="s_parD" style:family="table-cell" style:parent-style-name="Default"/>'
COMMON = ('<number:number-style style:name="N_common"><number:number number:decimal-places="2"/></number:number-style>'
          '<style:style style:name="Default" style:family="table-cell"/><style:style style:name="Fmt" style:family="table-cell" style:data-style-name="N_common"/>')
ONE = "<table:table-column/>"


def column(style):
    return f'<table:table-column table:default-cell-style-name="{style}"/>'


# name: (columns, attributes of the row, style of the cell, column of the cell)
PLACEMENT = {
    "bare (one column)": (ONE, "", None, 0),
    "no table:table-column": ("", "", None, 0),
    "past the declared columns": (ONE, "", None, 1),
    "second of two columns": (ONE + ONE, "", None, 1),
    "column default Default": (column("Default"), "", None, 0),
    "column default s_empty": (column("s_empty"), "", None, 0),
    "column default s_parD (parent Default)": (column("s_parD"), "", None, 0),
    "column default s_dec2": (column("s_dec2"), "", None, 0),
    "column default s_par (parent has N_common)": (column("s_par"), "", None, 0),
    "column default Fmt (N_common)": (column("Fmt"), "", None, 0),
    "column default that does not exist": (column("ghost"), "", None, 0),
    "row default s_dec2": (ONE, ' table:default-cell-style-name="s_dec2"', None, 0),
    "row default s_empty": (ONE, ' table:default-cell-style-name="s_empty"', None, 0),
    "row s_empty over column s_dec2": (column("s_dec2"), ' table:default-cell-style-name="s_empty"', None, 0),
    "row s_dec2 over column s_empty": (column("s_empty"), ' table:default-cell-style-name="s_dec2"', None, 0),
    "cell s_empty": (ONE, "", "s_empty", 0),
    "cell Default": (ONE, "", "Default", 0),
    "cell s_empty over column s_dec2": (column("s_dec2"), "", "s_empty", 0),
    "cell that does not exist": (ONE, "", "ghost", 0),
}
for k in list(NUMBERS) + list(OTHER):
    PLACEMENT[f"column default s_{k}"] = (column(f"s_{k}"), "", None, 0)
    PLACEMENT[f"cell s_{k}"] = (ONE, "", f"s_{k}", 0)


def cells(row_attributes, style, pad, kinds=TYPES):
    rows = []
    for _, f, attrs, _ in kinds:
        st = f' table:style-name="{style}"' if style else ""
        text = "S" if "string" in attrs else "999"
        lead = f'<table:table-cell table:number-columns-repeated="{pad}"/>' if pad else ""
        rows.append(f'<table:table-row{row_attributes}>{lead}<table:table-cell table:formula="{f}"{st} {attrs}><text:p>{text}</text:p></table:table-cell></table:table-row>')
    return "".join(rows)


def document(tables, version="1.3"):
    v = f' office:version="{version}"' if version else ""
    return (f'<?xml version="1.0" encoding="UTF-8"?><office:document-content {NS}{v}><office:automatic-styles>{AUTO}</office:automatic-styles>'
            f'<office:body><office:spreadsheet>{"".join(tables)}</office:spreadsheet></office:body></office:document-content>')


def package(path, content, generator, version="1.3"):
    meta = f'<meta:generator>{generator}</meta:generator>' if generator else ""
    with zipfile.ZipFile(path, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), "application/vnd.oasis.opendocument.spreadsheet")
        z.writestr("META-INF/manifest.xml", '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" '
                   f'manifest:version="{version or "1.3"}"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'
                   '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>'
                   '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/></manifest:manifest>')
        z.writestr("content.xml", content)
        z.writestr("styles.xml", f'<?xml version="1.0" encoding="UTF-8"?><office:document-styles {NS} office:version="{version or "1.3"}"><office:styles>{COMMON}</office:styles></office:document-styles>')
        z.writestr("meta.xml", f'<?xml version="1.0" encoding="UTF-8"?><office:document-meta {NS} office:version="{version or "1.3"}"><office:meta>{meta}</office:meta></office:document-meta>')


def matrix(attrs, f="of:=1+1", extra=""):
    return (f'<table:table-row><table:table-cell table:formula="{f}"{extra} table:number-matrix-columns-spanned="1" table:number-matrix-rows-spanned="1" {attrs}>'
            "<text:p>999</text:p></table:table-cell></table:table-row>")


EXTRA = {  # name: (rows, [(row, column, what the formula gives)])
    "array 1x1, float": (matrix('office:value-type="float" office:value="999"'), [(0, 0, 2)]),
    "array 1x1, string": (matrix('office:value-type="string" office:string-value="S"', 'of:=CONCATENATE(&quot;a&quot;;&quot;b&quot;)'), [(0, 0, "ab")]),
    "array 1x1, percentage": (matrix('office:value-type="percentage" office:value="9.99"'), [(0, 0, 2)]),
    "array 1x1, boolean": (matrix('office:value-type="boolean" office:boolean-value="false"', "of:=1=1"), [(0, 0, 1)]),
    "array 1x1, float with a style": (matrix('office:value-type="float" office:value="999"', extra=' table:style-name="s_dec2"'), [(0, 0, 2)]),
    "array 2x1 (master and covered cell)": (
        '<table:table-row><table:table-cell table:formula="of:=ROW([.B1:.B2])" table:number-matrix-columns-spanned="1" table:number-matrix-rows-spanned="2" '
        'office:value-type="float" office:value="999"><text:p>999</text:p></table:table-cell></table:table-row><table:table-row><table:covered-table-cell '
        'office:value-type="float" office:value="888"><text:p>888</text:p></table:covered-table-cell></table:table-row>', [(0, 0, 1), (1, 0, 2)]),
    "float repeated 3 columns": (
        '<table:table-row><table:table-cell table:formula="of:=1+1" table:number-columns-repeated="3" office:value-type="float" office:value="999"><text:p>999</text:p></table:table-cell></table:table-row>',
        [(0, 0, 2), (0, 1, 2), (0, 2, 2)]),
}


def generate(folder):
    os.makedirs(folder, exist_ok=True)
    bare = f'<table:table table:name="T">{ONE}{cells("", None, 0)}</table:table>'
    for (name, generator), version in itertools.product(GENERATORS.items(), ("1.3", "1.2", None)):
        # A document of 1.2 does not know the extension attributes (`calcext:value-type`): LibreOffice refuses to load it with them
        # (but for the one that marks an error).
        content = document([bare], version)
        if version != "1.3":
            for kind in ("float", "percentage", "currency", "boolean", "string", "date", "time"):
                content = content.replace(CX.format(kind), "")
        package(f"{folder}/generators_{name}_{version}.ods", content, generator, version)
    tables = [f'<table:table table:name="T{n}">{columns}{cells(row, style, pad)}</table:table>' for n, (columns, row, style, pad) in enumerate(PLACEMENT.values())]
    package(f"{folder}/placement.ods", document(tables), GENERATORS["lo26"])
    tables = [f'<table:table table:name="T{n}">{ONE}{rows}</table:table>' for n, (rows, _) in enumerate(EXTRA.values())]
    package(f"{folder}/extra.ods", document(tables), GENERATORS["lo26"])


def run(folder):
    import uno
    from com.sun.star.beans import PropertyValue

    def pv(n, v):
        p = PropertyValue()
        p.Name, p.Value = n, v
        return p

    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    def load(p):
        return desktop.loadComponentFromURL("file://" + os.path.abspath(p), "_blank", 0, (pv("Hidden", True),))

    def shows(cell, expected):
        return (cell.getString() == expected) if isinstance(expected, str) else abs(cell.getValue() - expected) < 1e-9

    head = f"{'':44}" + " ".join(f"{t[0][:8]:>8}" for t in TYPES)
    print("generators and versions (a bare cell of each type): R = calculated on load, . = kept")
    print(head)
    for p in sorted(glob.glob(folder + "/generators_*.ods")):
        d = load(p)
        sheet = d.Sheets.getByIndex(0)
        print(f"{os.path.basename(p)[11:-4]:44}" + " ".join(f"{'R' if shows(sheet.getCellByPosition(0, i), t[3]) else '.':>8}" for i, t in enumerate(TYPES)))
        d.close(True)
    d = load(folder + "/placement.ods")
    print("\nplacement of the style, and the type of the result")
    print(head)
    for n, (name, (_, _, _, pad)) in enumerate(PLACEMENT.items()):
        sheet = d.Sheets.getByIndex(n)
        print(f"{name:44}" + " ".join(f"{'R' if shows(sheet.getCellByPosition(pad, i), t[3]) else '.':>8}" for i, t in enumerate(TYPES)))
    d.close(True)
    d = load(folder + "/extra.ods")
    print("\narray formulas and repeated cells")
    for n, (name, (_, checks)) in enumerate(EXTRA.items()):
        sheet = d.Sheets.getByIndex(n)
        print(f"{name:44}" + " ".join("R" if shows(sheet.getCellByPosition(c, r), e) else "." for r, c, e in checks))
    d.close(True)


if __name__ == "__main__":
    {"gen": generate, "run": run}[sys.argv[1]](sys.argv[2])

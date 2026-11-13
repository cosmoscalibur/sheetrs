"""Makes `sheetrs-io/tests/assets/empty_text.{xlsx,ods}`: cells of sheet `T` that hold an empty text written
in every way a file can, and sheet `Calc`, whose row `i` has what LibreOffice 26.8 answers about `T!A<i>`
(`ISBLANK`, `COUNTA`, `LEN`, `TYPE`, ...), as the cached results of its formulas.

What LibreOffice says of each file is the oracle, and it differs per format: in an ODS a text cell that has
a paragraph, even an empty one, is an empty text (`ISBLANK` false, `COUNTA` 1), and a text cell with no
paragraph (`office:string-value=""` alone) is blank; in an XLSX every empty text of a shared string, an
inline string or a `t="str"` cell is a blank cell. A formula that gives an empty text is a formula with an
empty result in both.

The files are written by hand (LibreOffice would write a blank cell for each of them): a first pass of
each file has no cached results, LibreOffice opens it and recalculates, and the second pass writes the
file again with the results it gave for that very file. Run it with any Python and `soffice` on the PATH:

    python3 make_empty_text.py
"""
import os
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.dirname(HERE)
SOFFICE = os.environ.get("SOFFICE", "soffice")
NS = {
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "calcext": "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
}
q = lambda p, n: "{" + NS[p] + "}" + n

# What is asked about each cell of `T`, `{r}` standing for it.
PROBES = [
    "ISBLANK({r})", "COUNTA({r})", "COUNTBLANK({r})", "COUNT({r})", "ISTEXT({r})", "LEN({r})", "TYPE({r})",
    '{r}=""', "{r}=0", '{r}&"x"', 'COUNTIF({r},"")', 'COUNTIF({r},"=")', "ISNUMBER({r})", "ISNONTEXT({r})",
    "N({r})", "T({r})", 'IF(ISBLANK({r}),"b","nb")', '{r}&""',
]

# The cells of `T`: (what it is, the XML of the cell as the format writes it).
ODS_CELLS = [
    ("string, an empty paragraph", '<table:table-cell office:value-type="string"><text:p/></table:table-cell>'),
    ("string, string-value empty and an empty paragraph", '<table:table-cell office:value-type="string" office:string-value=""><text:p/></table:table-cell>'),
    ("string, string-value empty, no paragraph", '<table:table-cell office:value-type="string" office:string-value=""/>'),
    ("string, no value and no paragraph", '<table:table-cell office:value-type="string"/>'),
    ("string, a paragraph with start and end tags", '<table:table-cell office:value-type="string"><text:p></text:p></table:table-cell>'),
    ("string with calcext type, an empty paragraph", '<table:table-cell office:value-type="string" calcext:value-type="string"><text:p/></table:table-cell>'),
    ("no cell content", "<table:table-cell/>"),
    ('formula ="" with its empty result', '<table:table-cell table:formula="of:=&quot;&quot;" office:value-type="string" office:string-value="" calcext:value-type="string"><text:p/></table:table-cell>'),
    ('formula ="" with no result', '<table:table-cell table:formula="of:=&quot;&quot;"/>'),
    ("text abc", '<table:table-cell office:value-type="string" office:string-value="abc"><text:p>abc</text:p></table:table-cell>'),
    ("a space", '<table:table-cell office:value-type="string" office:string-value=" "><text:p> </text:p></table:table-cell>'),
]
XLSX_CELLS = [
    ("shared string, empty", ' t="s"', "<v>0</v>"),
    ("inline string, empty element", ' t="inlineStr"', "<is><t/></is>"),
    ("inline string, start and end tags", ' t="inlineStr"', "<is><t></t></is>"),
    ("t=str, empty value", ' t="str"', "<v></v>"),
    ("t=str, value element closed", ' t="str"', "<v/>"),
    ("t=str, no value", ' t="str"', ""),
    ("no cell content", "", ""),
    ('formula ="" with its empty result', ' t="str"', '<f>""</f><v></v>'),
    ('formula ="" with no result', ' t="str"', '<f>""</f>'),
    ("text abc", ' t="s"', "<v>1</v>"),
    ("a space", ' t="s"', "<v>2</v>"),
]


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def xlsx(path, results=None):
    rows = ""
    for i, (_, attrs, inner) in enumerate(XLSX_CELLS, 1):
        rows += f'<row r="{i}"><c r="A{i}"{attrs}>{inner}</c></row>'
    calc = ""
    for i in range(1, len(XLSX_CELLS) + 1):
        cs = ""
        for j, p in enumerate(PROBES):
            f = esc(p.format(r=f"T!A{i}"))
            kind, v = (results[i - 1][j] if results else (None, None))
            attr, val = "", ""
            if kind == "boolean":
                attr, val = ' t="b"', f"<v>{1 if v == 'true' else 0}</v>"
            elif kind == "float":
                val = f"<v>{v}</v>"
            elif kind == "string":
                attr, val = ' t="str"', f"<v>{esc(v)}</v>"
            elif kind is not None:
                attr, val = ' t="str"', "<v></v>"
            cs += f'<c r="{chr(65 + j)}{i}"{attr}><f>{f}</f>{val}</c>'
        calc += f'<row r="{i}">{cs}</row>'
    main = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    rel = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    rels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ct = "application/vnd.openxmlformats-officedocument.spreadsheetml"
    parts = {
        "[Content_Types].xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{ct}.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{ct}.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="{ct}.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="{ct}.sharedStrings+xml"/></Types>',
        "_rels/.rels": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rels}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook {main} {rel}><sheets><sheet name="T" sheetId="1" r:id="rId1"/><sheet name="Calc" sheetId="2" r:id="rId2"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rels}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{rels}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="{rels}/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
        "xl/sharedStrings.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst {main} count="3" uniqueCount="3"><si><t></t></si><si><t>abc</t></si><si><t xml:space="preserve"> </t></si></sst>',
        "xl/worksheets/sheet1.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet {main}><sheetData>{rows}</sheetData></worksheet>',
        "xl/worksheets/sheet2.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet {main}><sheetData>{calc}</sheetData></worksheet>',
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def ods(path, results=None):
    rows = "".join(f"<table:table-row>{c}</table:table-row>" for _, c in ODS_CELLS)
    calc = ""
    for i in range(1, len(ODS_CELLS) + 1):
        cs = ""
        for j, p in enumerate(PROBES):
            f = esc(p.format(r=f"[$T.A{i}]").replace(",", ";")).replace('"', "&quot;")
            kind, v = (results[i - 1][j] if results else (None, None))
            attrs = ""
            if kind == "boolean":
                attrs = f' office:value-type="boolean" office:boolean-value="{v}" calcext:value-type="boolean"'
            elif kind == "float":
                attrs = f' office:value-type="float" office:value="{v}" calcext:value-type="float"'
            elif kind == "string":
                attrs = f' office:value-type="string" office:string-value="{esc(v)}" calcext:value-type="string"'
            elif kind is not None:
                attrs = ' office:value-type="string" office:string-value="" calcext:value-type="string"'
            body = ""
            if kind == "string":
                body = f"<text:p>{esc(v)}</text:p>" if v else "<text:p/>"
            elif kind is not None:
                body = "<text:p/>"
            cs += f'<table:table-cell table:formula="of:={f}"{attrs}>{body}</table:table-cell>'
        calc += f"<table:table-row>{cs}</table:table-row>"
    content = (
        '<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
        'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" '
        f'office:version="1.3"><office:body><office:spreadsheet><table:table table:name="T">{rows}</table:table>'
        f'<table:table table:name="Calc">{calc}</table:table></office:spreadsheet></office:body></office:document-content>'
    )
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet", compress_type=zipfile.ZIP_STORED)
        z.writestr(
            "META-INF/manifest.xml",
            '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>',
            compress_type=zipfile.ZIP_DEFLATED,
        )
        z.writestr("content.xml", content, compress_type=zipfile.ZIP_DEFLATED)


def results_of(path, rows):
    """`[[(kind, value)]]` of `Calc` as LibreOffice 26.8 calculates it when it opens `path`."""
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "out")
        subprocess.run(
            [SOFFICE, f"-env:UserInstallation=file://{tmp}/profile", "--headless", "--convert-to", "ods", "--outdir", out, path],
            check=True, capture_output=True,
        )
        root = ET.fromstring(zipfile.ZipFile(os.path.join(out, os.path.splitext(os.path.basename(path))[0] + ".ods")).read("content.xml"))
    for sheet in root.iter(q("table", "table")):
        if sheet.get(q("table", "name")) != "Calc":
            continue
        res = []
        for row in sheet.iter(q("table", "table-row")):
            rep = int(row.get(q("table", "number-rows-repeated"), "1"))
            vals = []
            for c in row:
                if c.tag not in (q("table", "table-cell"), q("table", "covered-table-cell")):
                    continue
                k = c.get(q("calcext", "value-type")) or c.get(q("office", "value-type")) or "empty"
                if k == "boolean":
                    v = c.get(q("office", "boolean-value"))
                elif k == "float":
                    v = c.get(q("office", "value"))
                else:
                    v = "".join("".join(p.itertext()) for p in c.findall(q("text", "p")))
                vals += [(k, v)] * int(c.get(q("table", "number-columns-repeated"), "1"))
            res += [vals] * rep
        return res[:rows]
    raise SystemExit("no Calc sheet")


def main():
    for name, write, cells in (("xlsx", xlsx, XLSX_CELLS), ("ods", ods, ODS_CELLS)):
        path = os.path.join(OUT, f"empty_text.{name}")
        write(path)
        results = results_of(path, len(cells))
        write(path, results)
        fixture_determinism.normalize(path)
        print(path)


if __name__ == "__main__":
    main()

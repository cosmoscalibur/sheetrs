"""What LibreOffice 26.8 shows for a reference that went (`#REF!`) in the file saved and opened again, in both formats.

The reference of the project is the workbook SAVED AND OPENED AGAIN (`CONTRIBUTING.md`): each formula is written by hand in a
minimal XLSX (`#REF!` bare, `Sheet!#REF!`, `A1:#REF!`, a span of sheets, error constants, unknown names) and ODS (`[#REF!]`,
`[$Sheet.#REF!]`, `[$Sheet.A1:.#REF!]`), loaded, saved by LibreOffice and loaded again. Printed per formula: what it shows
on the first load (`first`: its memory, informational), after saving and loading (`reread`: the reference), and after a
hard recalculation of that file (`hardcalc`). Read the output with `docs/FIDELITY.md`, *A reference that went*.

    scripts/libreoffice-server.sh
    /opt/libreoffice26.8/program/python sheetrs-compiler/scripts/probe_lost_references.py [ods_bare]

The files it makes go to a temporary folder. The unit test that pins the result is
`a_reference_with_an_end_that_went_is_read_as_libreoffice_reads_it_saved` (`sheetrs-compiler/tests/semantics/pipeline.rs`).
"""
import os, sys, tempfile, zipfile, re
import uno, time
from com.sun.star.beans import PropertyValue
def pv(n,v):
    p=PropertyValue(); p.Name,p.Value=n,v; return p
def connect():
    local=uno.getComponentContext()
    r=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local)
    ctx=r.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    return ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
FILT={"xlsx":"Calc MS Excel 2007 XML","ods":"calc8"}
def new(desk):
    return desk.loadComponentFromURL("private:factory/scalc","_blank",0,(pv("Hidden",True),))
def load(desk,path):
    return desk.loadComponentFromURL("file://"+path,"_blank",0,(pv("Hidden",True),))
def save(doc,path,fmt):
    doc.storeToURL("file://"+path,(pv("FilterName",FILT[fmt]),))
def xml_formulas(path,fmt):
    z=zipfile.ZipFile(path); out=[]
    if fmt=="xlsx":
        for n in sorted(z.namelist()):
            if n.startswith("xl/worksheets/sheet"):
                s=z.read(n).decode()
                for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*>(?:<f([^>]*)>([^<]*)</f>|<f([^>]*)/>)(?:<v>([^<]*)</v>)?',s):
                    out.append((n[-10:],m.group(1),(m.group(3) if m.group(3) is not None else '(shared:'+(m.group(4) or '')+')'),m.group(5)))
        w=z.read("xl/workbook.xml").decode()
        for m in re.finditer(r'<definedName([^>]*)>([^<]*)</definedName>',w): out.append(("name",m.group(1),m.group(2),None))
    else:
        s=z.read("content.xml").decode()
        for m in re.finditer(r'<table:table-cell [^>]*table:formula="([^"]*)"[^>]*>',s):
            v=re.search(r'office:value="([^"]*)"',m.group(0)); t=re.search(r'office:string-value="([^"]*)"',m.group(0))
            out.append(("c",m.group(1),v.group(1) if v else (t.group(1) if t else None)))
        for m in re.finditer(r'<table:named-(?:range|expression) [^>]*>',s): out.append(("name",m.group(0)))
    return out

from xml.sax.saxutils import escape, quoteattr
def xlsx(path, formulas, shared=None, extra_sheet1="", raw=None):
    ct='''<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'''
    rels='''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
    wb='''<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/><sheet name="Calc" sheetId="3" r:id="rId3"/></sheets>%s</workbook>'''
    names=extra_sheet1
    wr='''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/></Relationships>'''
    ns='xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    s1=f'<?xml version="1.0" encoding="UTF-8"?><worksheet {ns}><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>1/0</f></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>x</t></is></c></row></sheetData></worksheet>'
    s2=f'<?xml version="1.0" encoding="UTF-8"?><worksheet {ns}><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>'
    rows=""
    for i,f in enumerate(formulas,1):
        if isinstance(f,tuple): addr,f=f
        else: addr=f"A{i}"
        r=int(''.join(c for c in addr if c.isdigit()))
        rows+=f'<row r="{r}"><c r="{addr}"><f>{escape(f)}</f></c></row>'
    s3=f'<?xml version="1.0" encoding="UTF-8"?><worksheet {ns}><sheetData>{rows}</sheetData></worksheet>'
    if raw: s3=raw
    with zipfile.ZipFile(path,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",ct); z.writestr("_rels/.rels",rels)
        z.writestr("xl/workbook.xml",wb%names); z.writestr("xl/_rels/workbook.xml.rels",wr)
        z.writestr("xl/worksheets/sheet1.xml",s1); z.writestr("xl/worksheets/sheet2.xml",s2); z.writestr("xl/worksheets/sheet3.xml",s3)
def ods(path, formulas, names=""):
    man='''<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>'''
    hdr='<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" office:version="1.3"><office:body><office:spreadsheet>'
    d='<table:table table:name="Data"><table:table-row><table:table-cell office:value-type="float" office:value="1"/><table:table-cell table:formula="of:=1/0"/></table:table-row><table:table-row><table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row></table:table>'
    o='<table:table table:name="Other"><table:table-row><table:table-cell office:value-type="float" office:value="7"/></table:table-row></table:table>'
    rows=""
    for f in formulas:
        rows+=f'<table:table-row><table:table-cell table:formula={quoteattr(f)}/></table:table-row>'
    c=f'<table:table table:name="Calc">{rows}</table:table>'
    content=hdr+d+o+c+names+'</office:spreadsheet></office:body></office:document-content>'
    with zipfile.ZipFile(path,"w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"),"application/vnd.oasis.opendocument.spreadsheet")
        z.writestr("META-INF/manifest.xml",man,zipfile.ZIP_DEFLATED); z.writestr("content.xml",content,zipfile.ZIP_DEFLATED)

H = tempfile.mkdtemp(prefix='lost-refs-')
d = connect()
CASES=[
 ("COUNT(#REF!)","COUNT([#REF!])"),("COUNTA(#REF!)","COUNTA([#REF!])"),("COUNTBLANK(#REF!)","COUNTBLANK([#REF!])"),
 ("COUNT(Data!#REF!)","COUNT([$Data.#REF!])"),("COUNTA(Data!#REF!)","COUNTA([$Data.#REF!])"),("COUNTBLANK(Data!#REF!)","COUNTBLANK([$Data.#REF!])"),
 ("COUNT(Data!A1:#REF!)","COUNT([$Data.A1:.#REF!])"),("COUNTA(Data!A1:#REF!)","COUNTA([$Data.A1:.#REF!])"),("COUNTBLANK(Data!A1:#REF!)","COUNTBLANK([$Data.A1:.#REF!])"),
 ("COUNTA(Data!#REF!:B2)","COUNTA([$Data.#REF!:.B2])"),("COUNTA(#REF!:B2)","COUNTA([#REF!:.B2])"),("COUNTA(A1:#REF!)","COUNTA([.A1:.#REF!])"),
 ("COUNT(Data:Other!#REF!)","COUNT([$Data.#REF!:$Other.#REF!])"),
 ("COUNT(Data!B1)","COUNT([$Data.B1])"),("COUNTA(Data!B1)","COUNTA([$Data.B1])"),("COUNTBLANK(Data!B1)","COUNTBLANK([$Data.B1])"),
 ("COUNTA(NA())","COUNTA(NA())"),("COUNT(NA())","COUNT(NA())"),("COUNTBLANK(NA())","COUNTBLANK(NA())"),
 ("COUNTA(1/0)","COUNTA(1/0)"),("COUNT(1/0)","COUNT(1/0)"),
 ("COUNT(Data!A1,#REF!)","COUNT([$Data.A1];[#REF!])"),("COUNTA(Data!A1,#REF!)","COUNTA([$Data.A1];[#REF!])"),
 ("COUNTA(#N/A)","COUNTA(#N/A)"),("COUNT(#N/A)","COUNT(#N/A)"),("COUNTA(#DIV/0!)","COUNTA(#DIV/0!)"),("COUNTA(#NAME?)","COUNTA(#NAME?)"),
 ("COUNTA(unknownname)","COUNTA(unknownname)"),("COUNT(unknownname)","COUNT(unknownname)"),
 ("#REF!+1","[#REF!]+1"),("ISERROR(#REF!)","ISERROR([#REF!])"),("ISREF(#REF!)","ISREF([#REF!])"),("ISNA(#REF!)","ISNA([#REF!])"),("IFERROR(#REF!,7)","IFERROR([#REF!];7)"),("SUM(#REF!)","SUM([#REF!])"),("ERROR.TYPE(#REF!)","ERROR.TYPE([#REF!])"),
 ("ISERROR(Data!#REF!)","ISERROR([$Data.#REF!])"),("ISREF(Data!#REF!)","ISREF([$Data.#REF!])"),("SUM(Data!#REF!)","SUM([$Data.#REF!])"),
 ("#REF!","[#REF!]"),("Data!#REF!","[$Data.#REF!]"),("COUNTA(#NAME?,1)","COUNTA(#NAME?;1)"),
 ("ROWS(#REF!)","ROWS([#REF!])"),("ROWS(Data!#REF!)","ROWS([$Data.#REF!])"),("ISBLANK(#REF!)","ISBLANK([#REF!])"),("ISNUMBER(#REF!)","ISNUMBER([#REF!])"),
]

CASES += [("ERROR.TYPE(#REF!)","ERROR.TYPE([#REF!])"),("ERROR.TYPE(unknownname)","ERROR.TYPE(unknownname)"),("ERROR.TYPE(#DIV/0!)","ERROR.TYPE(#DIV/0!)"),
("COUNTBLANK(unknownname)","COUNTBLANK(unknownname)"),("COUNTBLANK(#DIV/0!)","COUNTBLANK(#DIV/0!)"),("#DIV/0!+1","#DIV/0!+1"),("unknownname+1","unknownname+1"),
("SUM(Data!A1:unknownname)","SUM([$Data.A1]:unknownname)"),("SUM(Data!A1:#DIV/0!)","SUM([$Data.A1]:#DIV/0!)"),("COUNTA(Data!A1:unknownname)","COUNTA([$Data.A1]:unknownname)")]

if len(sys.argv) > 1 and sys.argv[1] == "ods_bare":
    CASES = [(a, a.replace("[", "").replace("]", "")) for a, b in CASES]


def readback(doc, n):
    sheet = doc.Sheets.getByName("Calc")
    return [sheet.getCellByPosition(0, i).String for i in range(n)]


for fmt in ("xlsx", "ods"):
    formulas = [c[0] if fmt == "xlsx" else c[1] for c in CASES]
    path = f"{H}/h.{fmt}"
    (xlsx if fmt == "xlsx" else ods)(path, formulas)
    doc = load(d, path)
    first = readback(doc, len(formulas))
    own = f"{H}/h_rs.{fmt}"
    save(doc, own, fmt)
    doc.calculateAll()
    hard = readback(doc, len(formulas))
    doc.close(True)
    doc = load(d, own)
    reread = readback(doc, len(formulas))
    doc.close(True)
    print("=====", fmt)
    for i, f in enumerate(formulas):
        print(f"{f:36s} first={first[i]!r:10} reread={reread[i]!r:10} hardcalc={hard[i]!r}")

"""Probe: what LibreOffice 26.8 shows of a whole column over several sheets (`SUM($Data.A1:$Wide.A1048576)`) in the
file it saves and opens again (the reference of the project, `CONTRIBUTING.md`), in an XLSX and in an ODS: the default
load keeps the cached sum, and only a hard recalculation (`calculateAll`) of the XLSX gives `#NAME?` (its own form,
`SUM(Data:Wide!A:A)`, is not read back by its parser). It is why this is no difference of SheetRS
(`sheetrs-compiler/docs/FIDELITY.md`).

    scripts/libreoffice-server.sh
    /opt/libreoffice26.8/program/python sheetrs-compiler/scripts/probe_reread_whole_columns.py
"""
import os, tempfile
import uno
from com.sun.star.beans import PropertyValue
def pv(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
local=uno.getComponentContext()
resolver=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
ctx=resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
desktop=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
doc=desktop.loadComponentFromURL("private:factory/scalc","_blank",0,(pv("Hidden",True),))
doc.Sheets.insertNewByName("Data",0); doc.Sheets.insertNewByName("Wide",1); doc.Sheets.insertNewByName("Calc",2)
for i,(sh,v) in enumerate([("Data",1),("Data",2),("Wide",10),("Wide",20)]):
    doc.Sheets.getByName(sh).getCellByPosition(0,i%2).setValue(v)
c=doc.Sheets.getByName("Calc")
forms=["=SUM($Data.A1:$Wide.A1048576)","=SUM($Data.A:$Wide.A)","=SUM($Data.A1:$Wide.A3)","=COUNT($Data.A1:$Wide.A1048576)"]
for i,f in enumerate(forms): c.getCellByPosition(0,i).setFormula(f)
def show(d,tag):
    s=d.Sheets.getByName("Calc")
    print(tag,[(s.getCellByPosition(0,i).getFormula(), s.getCellByPosition(0,i).getString()) for i in range(len(forms))])
show(doc,"memory")
with tempfile.TemporaryDirectory() as f:
    for ext,flt in (("xlsx","Calc MS Excel 2007 XML"),("ods","calc8")):
        p=os.path.join(f,"o."+ext)
        doc.storeToURL("file://"+p,(pv("FilterName",flt),))
        d2=desktop.loadComponentFromURL("file://"+p,"_blank",0,(pv("Hidden",True),))
        show(d2,ext+" reread")
        d2.calculateAll(); show(d2,ext+" reread+calculateAll")
        d2.close(True)
        if ext=="xlsx":
            import zipfile
            z=zipfile.ZipFile(p); print([l for l in z.read("xl/worksheets/sheet3.xml").decode().split("<f") ][1:5])
doc.close(True)

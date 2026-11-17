"""What LibreOffice 26.8 calculates again when a row is hidden (a probe, not a test): the formulas are `SUM` over ranges of `A1:A6`
(and one `SUBTOTAL`) with a stale result cached (the text `S`); after hiding rows, saving and opening again, the formulas that show
something else were calculated.

    python3 probe_hide.py gen [adjacent]                              # writes the workbook (needs openpyxl), spacing the formulas
    /opt/libreoffice26.8/program/python probe_hide.py run [adjacent]  # LibreOffice listening on port 2002

With the formulas in adjacent cells (`D1:D17`) LibreOffice also calculates `SUM(A1:A3)` and `SUM(A2:A4)` when row 2 is hidden (formulas of one
shape that sit together are a group, which listens to the cells it reads as one area); with two empty rows between them (`D1, D4, D7...`,
as here) only `SUBTOTAL` is calculated, the rule `SheetRS` follows (`docs/approved-differences.md`, *Proposed*).
"""
import os,re,sys,zipfile,subprocess,tempfile
import tempfile
P=os.path.join(tempfile.gettempdir(),"sheetrs-probe-hide.xlsx")
STEP = 1 if "adjacent" in sys.argv else 3  # `adjacent`: the formulas in D1:D17
F=["=SUM(A2:A3)","=SUM(A2:A2)","=SUM(A3:A4)","=SUM(A1:A3)","=SUM(A2:A4)","=SUM(B2:B3)","=A2","=A3","=SUM(A1:A6)","=SUM(A3:A5)","=SUBTOTAL(109,A2:A3)","=SUM(A2:B3)","=COUNT(A2:A3)","=MAX(A2:A3)","=SUM(A4:A5)","=SUM(A5:A6)","=SUM(A1:A2)"]
if sys.argv[1]=="gen":
    import openpyxl
    wb=openpyxl.Workbook(); s=wb.active
    for i in range(1,7): s.cell(row=i,column=1,value=i*10); s.cell(row=i,column=2,value=i)
    for k,f in enumerate(F,1): s[f"D{STEP*k-(STEP-1)}"]=f
    wb.save(P)
    tmp=P+".tmp"
    with zipfile.ZipFile(P) as zin, zipfile.ZipFile(tmp,"w",zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            b=zin.read(it.filename)
            if it.filename=="xl/worksheets/sheet1.xml":
                x,n=re.subn(r'<c r="(D\d+)"([^>]*)><f>(.*?)</f>(?:<v\s*/>|<v></v>)</c>',lambda m:'<c r="%s"%s t="str"><f>%s</f><v>S</v></c>'%(m.group(1),m.group(2),m.group(3)),b.decode())
                print("patched",n); b=x.encode()
            if it.filename=="xl/workbook.xml": b=b.decode().replace(' fullCalcOnLoad="1"',"").encode()
            zout.writestr(it,b)
    os.replace(tmp,P)
    sys.exit()
import uno
from com.sun.star.beans import PropertyValue
def pv(n,v):
    p=PropertyValue();p.Name,p.Value=n,v;return p
local=uno.getComponentContext()
r=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local)
ctx=r.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
desk=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
def load(p): return desk.loadComponentFromURL("file://"+p,"_blank",0,(pv("Hidden",True),))
for hide in ([2],[3],[1],[2,3],[5],[6]):
    doc=load(P)
    for h in hide: doc.Sheets.getByIndex(0).Rows.getByIndex(h-1).IsVisible=False
    with tempfile.TemporaryDirectory() as t:
        o=t+"/x.xlsx"; doc.storeToURL("file://"+o,(pv("FilterName","Calc MS Excel 2007 XML"),)); doc.close(True)
        d2=load(o); s2=d2.Sheets.getByIndex(0)
        print("hide rows",hide,[ F[k] for k in range(len(F)) if s2.getCellByPosition(3,STEP*k).getString()!="S"]); d2.close(True)

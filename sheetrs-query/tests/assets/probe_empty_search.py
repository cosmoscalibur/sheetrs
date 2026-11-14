"""Probes the empty search string with LibreOffice: `findAll` and `replaceAll` over `anchors.xlsx` and
`anchors.ods`, in every mode, with `entire_cells` on and off and an empty and a non-empty
replacement. LibreOffice 26.8 finds no cell (`findAll` gives nothing) and changes none, in all of them:
the empty search text matches nothing. `sheetrs_protocol::Find::matches_nothing` answers it at once.

Run it with a LibreOffice listening on a socket (`scripts/libreoffice-server.sh`):

    /opt/libreoffice26.8/program/python probe_empty_search.py
"""
import os, time
import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def set_mode(d, mode, entire):
    d.SearchString = ""
    d.SearchRegularExpression = mode == "regex"
    d.SearchWildcard = mode == "wildcards"
    d.SearchWords = entire


for ext in ("xlsx", "ods"):
    doc = desktop.loadComponentFromURL("file://" + os.path.join(HERE, f"anchors.{ext}"), "_blank", 0, (pv("Hidden", True),))
    sheet = doc.Sheets.getByIndex(0)
    before = [sheet.getCellByPosition(0, i).String for i in range(21)]
    for mode in ("literal", "regex", "wildcards"):
        for entire in (False, True):
            d = sheet.createSearchDescriptor()
            set_mode(d, mode, entire)
            assert sheet.findAll(d) is None, (ext, mode, entire)
            for with_ in ("", "Z"):
                d = sheet.createReplaceDescriptor()
                set_mode(d, mode, entire)
                d.ReplaceString = with_
                assert sheet.replaceAll(d) == 0, (ext, mode, entire, with_)
                assert [sheet.getCellByPosition(0, i).String for i in range(21)] == before
    doc.close(True)
print("the empty search string finds nothing and replaces nothing, in every mode")

import random, zipfile, struct, sys
random.seed(20261116)
N_PER = int(sys.argv[1]) if len(sys.argv) > 1 else 25000
def families():
    out = []
    out += [("money", random.randrange(0, 10**9) / 100) for _ in range(N_PER)]
    for _ in range(N_PER):
        d = random.randrange(0, 6); n = random.randrange(0, 10**6)
        out.append(("tie", n / 10**d + 5 / 10**(d + 1)))
    out += [("real", random.uniform(0, 1000)) for _ in range(N_PER)]
    for _ in range(N_PER):
        out.append(("calc", round(random.uniform(1, 500), 2) * random.randrange(1, 100) * round(random.uniform(0.5, 2.0), 4)))
    out += [("wide", random.uniform(1, 10) * 10 ** random.randrange(-8, 13)) for _ in range(N_PER)]
    for _ in range(N_PER):
        out.append(("noise", sum(round(random.uniform(0, 100), 2) for _ in range(random.randrange(2, 6))) / random.choice([1, 3, 7, 12])))
    res = []
    for fam, x in out:
        res.append((fam, x)); res.append((fam + "-", -x))
    return res
cases = families()
ds = [random.choice([0, 0, 1, 2, 2, 2, 3, 4, 6, -1, -2, 8]) for _ in cases]
rows = []
for i, ((fam, x), d) in enumerate(zip(cases, ds), 1):
    rows.append(f'<row r="{i}"><c r="A{i}"><v>{x!r}</v></c><c r="B{i}"><v>{d}</v></c>'
                f'<c r="C{i}"><f>ROUND(A{i},B{i})</f></c><c r="D{i}"><f>ROUNDUP(A{i},B{i})</f></c><c r="E{i}"><f>ROUNDDOWN(A{i},B{i})</f></c></row>')
sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + "".join(rows) + '</sheetData></worksheet>'
with zipfile.ZipFile("cases_in.xlsx", "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    z.writestr("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>')
    z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
    z.writestr("xl/worksheets/sheet1.xml", sheet)
with open("meta.tsv", "w") as f:
    for (fam, x), d in zip(cases, ds): f.write(f"{fam}\t{x!r}\t{d}\n")
print(len(cases), "rows")

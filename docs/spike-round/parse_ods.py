"""Reads what LibreOffice calculated (the ODS it saved) and writes `cases.tsv` for compare.rs: family, the
input as LibreOffice holds it (hex of the f64), digits, function (0 ROUND, 1 ROUNDUP, 2 ROUNDDOWN) and its
result (hex). The input is the one LibreOffice read, not the one written: its importer keeps 15 significant
digits of a longer literal.

    python3 parse_ods.py meta.tsv out/cases_in.ods cases.tsv
"""
import re, struct, sys, zipfile

meta = [l.rstrip("\n").split("\t") for l in open(sys.argv[1])]
xml = zipfile.ZipFile(sys.argv[2]).read("content.xml").decode()
rows = re.findall(r"<table:table-row[^>]*>(.*?)</table:table-row>", xml, re.S)
cell = re.compile(r"<table:table-cell([^>]*?)(?:/>|>)")
n = 0
with open(sys.argv[3], "w") as out:
    for (fam, _, d), row in zip(meta, rows):
        vals = []
        for m in cell.finditer(row):
            a = m.group(1)
            rep = re.search(r'number-columns-repeated="(\d+)"', a)
            v = re.search(r'office:value="([^"]*)"', a)
            vals += [v.group(1) if v else None] * (int(rep.group(1)) if rep else 1)
        x = float(vals[0])
        for fn in range(3):
            if vals[2 + fn] is not None:
                out.write(f"{fam}\t{struct.pack('>d', x).hex()}\t{d}\t{fn}\t{struct.pack('>d', float(vals[2 + fn])).hex()}\n")
                n += 1
print(n, "cases")

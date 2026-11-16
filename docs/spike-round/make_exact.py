import random, zipfile, math, sys
random.seed(20261117)
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
rows = []; meta = []
for i, ((fam, x), d) in enumerate(zip(cases, ds), 1):
    m, e = math.frexp(abs(x)); M = int(m * 2**53); E = e - 53
    assert M * 2.0**E == abs(x)
    a, b = M >> 26, M & (2**26 - 1); s = -1 if x < 0 else 1
    rows.append(f'<row r="{i}"><c r="A{i}"><v>{a}</v></c><c r="B{i}"><v>{b}</v></c><c r="C{i}"><v>{E}</v></c><c r="D{i}"><v>{d}</v></c><c r="E{i}"><v>{s}</v></c>'
                f'<c r="F{i}"><f>E{i}*(A{i}*67108864+B{i})*POWER(2,C{i})</f></c>'
                f'<c r="G{i}"><f>ROUND(F{i},D{i})</f></c><c r="H{i}"><f>ROUNDUP(F{i},D{i})</f></c><c r="I{i}"><f>ROUNDDOWN(F{i},D{i})</f></c></row>')
    meta.append(f"{fam}\t{x!r}\t{d}")
sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + "".join(rows) + '</sheetData></worksheet>'
src = zipfile.ZipFile("cases_in.xlsx")
with zipfile.ZipFile("cases_exact.xlsx", "w", zipfile.ZIP_DEFLATED) as z:
    for n in src.namelist(): z.writestr(n, sheet if n == "xl/worksheets/sheet1.xml" else src.read(n))
open("meta_exact.tsv", "w").write("\n".join(meta) + "\n")
print(len(cases), "rows")

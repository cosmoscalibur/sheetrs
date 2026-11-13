#!/usr/bin/env python3
"""What LibreOffice and the engine answer to the same formulas, side by side.

    scripts/lo_vs_engine.py FORMULAS.txt        # one formula per line, XLSX syntax, no `=`

The formulas go to the rows of the `Calc` sheet of the oracle workbook (`gen_semantics_fixture.build`, so
`Data!A1:A8`, the long ranges of `Wide` and the rest are the ones the oracle uses); LibreOffice 26.8
recalculates it (`lo_probe.py`) and the engine runs it (`sheetc calculate`, `$SHEETC`, by default
`target/release/sheetc`). A line starts with `!!` when the answers differ. A formula the engine cannot
run (syntax it refuses, a count of arguments it does not take) fails the whole file: it says which.
Rows that agree go to `scripts/*_formulas.txt` to become rows of the oracle.
"""
import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import gen_semantics_fixture as gen  # noqa: E402
import lo_probe  # noqa: E402
import probe_intersection as pi  # noqa: E402

SHEETC = os.environ.get("SHEETC", os.path.join(HERE, "..", "..", "target", "release", "sheetc"))
fs=[l.rstrip('\n') for l in open(sys.argv[1],encoding='utf-8') if l.strip() and not l.startswith('#')]
lo=lo_probe.probe_all(fs)
with tempfile.TemporaryDirectory() as tmp:
    p=os.path.join(tmp,'e.xlsx'); gen.build(p,fs)
    r=subprocess.run([SHEETC,'calculate','-i',p],input='',capture_output=True,text=True)
if '{' not in r.stdout:
    print('ENGINE FAIL:',(r.stdout+r.stderr)[-1500:]); eng=[None]*len(fs)
else:
    answer=json.loads(r.stdout[r.stdout.index('{'):])['valueRanges']
    rows=next(e['values'] for e in answer if e['name'].startswith('Calc!'))
    eng=[pi.efmt(row[0]) for row in rows]
bad=0
for f,(k,v),e in zip(fs,lo,eng):
    ok = e is not None and pi.same(v,e)
    if not ok: bad+=1
    print(('  ' if ok else '!!')+f"{f}\tLO={k}:{v}\tENG={e}")
print('differ',bad,'of',len(fs))

#!/usr/bin/env python3
"""Synthetic large workbook for performance work: ~800k formulas in the shapes of
real models (row-wise arithmetic, IF, exact VLOOKUP into a table, running totals,
SUMIF over blocks, text assembly), one declared input and a few outputs.

Usage: gen_large_fixture.py OUT.xlsx [ROWS]   (default 100000 rows of Calc)
"""
import random
import sys

import xlsxwriter

out = sys.argv[1]
rows = int(sys.argv[2]) if len(sys.argv) > 2 else 100_000
random.seed(7)
wb = xlsxwriter.Workbook(out, {"constant_memory": False})
data = wb.add_worksheet("Data")
lookup = wb.add_worksheet("Lookup")
calc = wb.add_worksheet("Calc")
summ = wb.add_worksheet("Summary")
sett = wb.add_worksheet("settings")

keys = [f"K{i:04d}" for i in range(2000)]
for i, k in enumerate(keys):
    lookup.write(i, 0, k)
    lookup.write(i, 1, round(random.uniform(0.5, 2.0), 4))
data.write(0, 7, "rate")
data.write(0, 8, 0.05)                      # I1: the input
for r in range(rows):
    data.write(r + 1, 0, random.randint(1, 1000))
    data.write(r + 1, 1, round(random.uniform(1, 500), 2))
    data.write(r + 1, 2, random.choice(keys))
    data.write(r + 1, 3, random.randint(0, 11))
for r in range(1, rows + 1):
    n = r + 1
    calc.write_formula(r, 0, f"=Data!A{n}*Data!B{n}")
    calc.write_formula(r, 1, f"=IF(A{n}>50000,A{n}*(1-Data!$I$1),A{n})")
    calc.write_formula(r, 2, f"=VLOOKUP(Data!C{n},Lookup!$A$1:$B$2000,2,FALSE)")
    calc.write_formula(r, 3, f"=ROUND(B{n}*C{n},2)")
    calc.write_formula(r, 4, f"=D{n}" if r == 1 else f"=E{n-1}+D{n}")
    calc.write_formula(r, 5, f'=Data!C{n}&"-"&TEXT(Data!D{n},"00")')
    calc.write_formula(r, 6, f"=IFERROR(D{n}/Data!A{n},0)")
    calc.write_formula(r, 7, f"=MAX(D{n},G{n})-MIN(D{n},G{n})")
summ.write_formula(0, 0, f"=SUM(Calc!D2:D{rows+1})")
summ.write_formula(1, 0, f"=Calc!E{rows+1}")
summ.write_formula(2, 0, f'=SUMIF(Data!D2:D{rows+1},3,Calc!D2:D{rows+1})')
summ.write_formula(3, 0, f"=COUNTIF(Calc!F2:F{rows+1},\"K00*\")")
summ.write_formula(4, 0, f"=AVERAGE(Calc!H2:H{rows+1})")
for c, h in enumerate(["key", "value", "description", "meta_type"]):
    sett.write(0, c, h)
for i, row in enumerate([
    ("rate", "Data!I1", "Discount rate", "input"),
    ("total", "Summary!A1", "Sum of D", "output"),
    ("running", "Summary!A2", "Final running total", "output"),
    ("month3", "Summary!A3", "SUMIF month 3", "output"),
    ("keys", "Summary!A4", "COUNTIF prefix", "output"),
    ("spread", "Summary!A5", "Average spread", "output"),
]):
    for c, v in enumerate(row):
        sett.write(i + 1, c, v)
wb.close()
print(out, rows * 8, "formulas")

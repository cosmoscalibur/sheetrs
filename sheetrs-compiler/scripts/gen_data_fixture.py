#!/usr/bin/env python3
"""Large data workbook for performance work: 600,000 rows of 8 random numbers, a label, a
key and a row `SUM`, plus a Summary sheet with a total, a `COUNTIF`, an exact `VLOOKUP` and
an input. About 49 MB on disk, 290 MB of XML, 600,005 formulas, no cached values.

Usage: gen_data_fixture.py OUT.xlsx [--declared]       (random.seed(1); same data every run, the zip bytes differ with the clock)

--declared adds a `settings` sheet (input `scale`, outputs `total`, `alpha_count`, `lookup`, `scaled`),
so a build keeps only what those outputs need instead of every cell.
"""
import random
import sys

import openpyxl

out = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1] != "--declared" else "big.xlsx"
declared = "--declared" in sys.argv
wb = openpyxl.Workbook(write_only=True)
ws = wb.create_sheet("Data")
random.seed(1)
N = 600_000
labels = ["alpha", "beta", "gamma", "delta", "épsilon", "ñandú"]
for r in range(1, N + 1):
    row = [round(random.random() * 1000, 3) for _ in range(8)] + [labels[r % 6], r]
    row.append(f"=SUM(A{r}:H{r})")
    ws.append(row)
s = wb.create_sheet("Summary")
s.append(["Total", "=SUM(Data!K1:K600000)"])
s.append(["Alpha count", '=COUNTIF(Data!I1:I600000,"alpha")'])
s.append(["Lookup", "=VLOOKUP(Data!J300000,Data!J1:K600000,2,FALSE)"])
s.append(["Input", 1])
s.append(["Scaled", "=B4*B1"])
if declared:
    c = wb.create_sheet("settings")
    c.append(["key", "value", "description", "meta_type"])
    c.append(["scale", "Summary!B4", "Multiplier of the total", "input"])
    c.append(["total", "Summary!B1", "Sum of the row sums", "output"])
    c.append(["alpha_count", "Summary!B2", "Rows labelled alpha", "output"])
    c.append(["lookup", "Summary!B3", "Row sum of key 300000", "output"])
    c.append(["scaled", "Summary!B5", "Total times scale", "output"])
wb.save(out)

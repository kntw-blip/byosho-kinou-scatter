"""Dump sheet names and the first rows of each downloaded workbook (for inspecting layout)."""
import json, sys, glob, zipfile
from openpyxl import load_workbook

out = {}
for path in sorted(glob.glob("raw/*.xlsx")):
    wb = load_workbook(path, read_only=True, data_only=True)
    info = {}
    for ws in wb.worksheets:
        rows = []
        n = 0
        for row in ws.iter_rows(values_only=True):
            if n < 12:
                rows.append([None if v is None else str(v) for v in row])
            n += 1
        info[ws.title] = {"n_rows": n, "head": rows}
    out[path] = info
for path in glob.glob("raw/*.zip"):
    out[path] = zipfile.ZipFile(path).namelist()
json.dump(out, open("probe.json", "w"), ensure_ascii=False)

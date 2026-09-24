"""
inspect_dispatch_xls.py
========================
Inspects the sheets and columns of:
  1. d:\Alpha\Tubex Records\dispatch nov to jul.xls
  2. d:\Alpha\Tubex Records\dispatch pet nov to jul.xls
"""

import os, xlrd

RECORDS_DIR = r"d:\Alpha\Tubex Records"

def inspect_file(filename):
    path = os.path.join(RECORDS_DIR, filename)
    if not os.path.exists(path):
        print(f"File not found: {path}")
        return
    print("\n" + "="*60)
    print(f" FILE: {filename}")
    print("="*60)
    wb = xlrd.open_workbook(path)
    print(f" Sheet names: {wb.sheet_names()}")
    for s_idx, s_name in enumerate(wb.sheet_names()):
        ws = wb.sheet_by_index(s_idx)
        print(f"\n --- Sheet [{s_name}] ({ws.nrows} rows, {ws.ncols} cols) ---")
        for r in range(min(15, ws.nrows)):
            row = [ws.cell_value(r, c) for c in range(ws.ncols)]
            print(f" Row {r:2d}: {row[:10]}")

inspect_file("dispatch nov to jul.xls")
inspect_file("dispatch pet nov to jul.xls")

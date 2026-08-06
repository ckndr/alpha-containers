"""
inspect_xlsx_dispatch.py
========================
Inspects Dispatch sheets in Tubex_Aug26.xlsx and Tubex_July26.xlsx
"""

import os, openpyxl

ALPHA_DIR   = r"d:\Alpha"
RECORDS_DIR = r"d:\Alpha\Tubex Records"

def inspect_wb(path, label):
    if not os.path.exists(path):
        print(f"{label} not found: {path}")
        return
    print("\n" + "="*60)
    print(f" {label}: {os.path.basename(path)}")
    print("="*60)
    wb = openpyxl.load_workbook(path, data_only=True)
    print(" Sheet names:", wb.sheetnames)
    for sname in ["Dispatch", "Dispatch_Log", "Dispatch_Summary", "Dispatch Log"]:
        if sname in wb.sheetnames:
            ws = wb[sname]
            print(f"\n --- Sheet [{sname}] ({ws.max_row} rows, {ws.max_column} cols) ---")
            for r in range(1, min(15, ws.max_row + 1)):
                row = [ws.cell(r, c).value for c in range(1, min(12, ws.max_column + 1))]
                print(f" Row {r:2d}: {row}")
    wb.close()

inspect_wb(os.path.join(ALPHA_DIR, "Tubex_Aug26.xlsx"), "ACTIVE FILE")
inspect_wb(os.path.join(RECORDS_DIR, "Tubex_July26.xlsx"), "JULY ARCHIVE FILE")

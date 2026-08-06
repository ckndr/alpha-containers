"""
inspect_july_daily_rows.py
==========================
Inspects line-by-line rows in:
  1. d:\Alpha\Tubex Records\Tubex_July26.xlsx (Production_Log)
  2. d:\Alpha\Tubex Records\production nov to jul.xls (July rows)
"""

import os, xlrd, openpyxl, datetime

RECORDS_DIR = r"d:\Alpha\Tubex Records"

print("="*70)
print(" 1. Tubex_July26.xlsx - Production_Log Sheet Rows")
print("="*70)
july_path = os.path.join(RECORDS_DIR, "Tubex_July26.xlsx")
wb = openpyxl.load_workbook(july_path, data_only=True)
ws = wb["Production_Log"]
july_log_rows = []
for r in range(3, ws.max_row + 1):
    row = [ws.cell(r, c).value for c in range(1, 11)]
    if row[0] and row[3]:
        date_val, machine, customer, product, dia, pid, target, good, reject, waste = (row + [None]*10)[:10]
        if good is not None:
            july_log_rows.append({
                "row_idx": r,
                "date": date_val,
                "customer": customer,
                "product": product,
                "dia": dia,
                "good": good,
                "reject": reject
            })

print(f"Total valid Production_Log rows in Tubex_July26.xlsx: {len(july_log_rows)}")
tube_tot = sum(r['good'] for r in july_log_rows if "ml" not in str(r['dia']).lower())
pet_tot  = sum(r['good'] for r in july_log_rows if "ml" in str(r['dia']).lower())
print(f"  TUBE Good Sum: {tube_tot:,}")
print(f"  PET  Good Sum: {pet_tot:,}")
print("  Sample rows:")
for r in july_log_rows[:5]:
    print("   ", r)
wb.close()

print("\n" + "="*70)
print(" 2. production nov to jul.xls - July Rows")
print("="*70)
leg_path = os.path.join(RECORDS_DIR, "production nov to jul.xls")
wb = xlrd.open_workbook(leg_path)
ws = wb.sheet_by_index(0)
leg_july = []
for r in range(10, ws.nrows):
    row = [ws.cell_value(r, c) for c in range(ws.ncols)]
    if not row or len(row) < 10: continue
    ref_no, date_val, line_no, shift, ptime, pof, pname, target, dia, good, wastage = (row + [None]*15)[:11]
    if pname and good not in (None, ""):
        dt = None
        if isinstance(date_val, float):
            try: dt = xlrd.xldate_as_datetime(date_val, wb.datemode)
            except: pass
        mstr = dt.strftime("%B %Y") if dt else "Unknown"
        if mstr == "July 2026":
            leg_july.append({
                "row_idx": r,
                "date": dt.strftime("%Y-%m-%d") if dt else date_val,
                "product": pname,
                "good": good,
                "reject": wastage
            })

print(f"Total July rows in legacy production nov to jul.xls: {len(leg_july)}")
leg_tube_tot = sum(r['good'] for r in leg_july)
print(f"  Legacy July TUBE Good Sum: {int(leg_tube_tot):,}")
print("  Legacy July rows:")
for r in leg_july:
    print("   ", r)

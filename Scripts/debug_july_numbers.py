"""
debug_july_numbers.py
=====================
Inspects July 2026 numbers across all sources:
  1. Tubex_July26.xlsx (Tubex_Dashboard & Production_Log)
  2. production nov to jul.xls (Legacy Production for July 2026)
  3. Production_Archive.xlsx (July 2026 tab & All Months tab)
  4. Tubex.html CUSTOMER_REPORT_DATA
"""

import os, xlrd, openpyxl, json, re

RECORDS_DIR = r"d:\Alpha\Tubex Records"
ALPHA_DIR   = r"d:\Alpha"

print("="*70)
print(" 1. INSPECTING Tubex_July26.xlsx ")
print("="*70)
july_path = os.path.join(RECORDS_DIR, "Tubex_July26.xlsx")
if os.path.exists(july_path):
    wb = openpyxl.load_workbook(july_path, data_only=True)
    print("Sheets:", wb.sheetnames)
    
    # Tubex_Dashboard KPI
    if "Tubex_Dashboard" in wb.sheetnames:
        ws_dash = wb["Tubex_Dashboard"]
        tube_mtd = ws_dash.cell(6, 4).value
        pet_mtd  = ws_dash.cell(8, 4).value
        tube_disp = ws_dash.cell(6, 10).value
        pet_disp  = ws_dash.cell(8, 10).value
        print(f" Tubex_Dashboard sheet KPIs:")
        print(f"   TUBE MTD Produced : {tube_mtd}")
        print(f"   TUBE MTD Dispatched: {tube_disp}")
        print(f"   PET MTD Produced  : {pet_mtd}")
        print(f"   PET MTD Dispatched : {pet_disp}")
        
    # Production_Log sum
    if "Production_Log" in wb.sheetnames:
        ws_log = wb["Production_Log"]
        tube_sum = 0
        pet_sum = 0
        rows_count = 0
        for r in range(3, ws_log.max_row + 1):
            date_val = ws_log.cell(r, 1).value
            pname    = ws_log.cell(r, 4).value
            dia      = str(ws_log.cell(r, 5).value or "")
            good     = ws_log.cell(r, 8).value
            if date_val and pname and good is not None:
                rows_count += 1
                qty = int(good or 0)
                if "ml" in dia.lower():
                    pet_sum += qty
                else:
                    tube_sum += qty
        print(f" Production_Log sheet Sum ({rows_count} rows):")
        print(f"   TUBE Produced sum : {tube_sum:,}")
        print(f"   PET Produced sum  : {pet_sum:,}")
    wb.close()

print("\n" + "="*70)
print(" 2. INSPECTING legacy production nov to jul.xls FOR July 2026")
print("="*70)
leg_path = os.path.join(RECORDS_DIR, "production nov to jul.xls")
if os.path.exists(leg_path):
    wb = xlrd.open_workbook(leg_path)
    ws = wb.sheet_by_index(0)
    july_tube_leg = 0
    july_pet_leg  = 0
    july_rows_leg = 0
    
    all_months_prod = {}
    
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
            qty = int(float(good))
            is_pet = ("ml" in str(dia).lower() or "PET" in str(pname).upper() or "BOTTLE" in str(pname).upper())
            
            if mstr not in all_months_prod:
                all_months_prod[mstr] = {"TUBE": 0, "PET": 0, "rows": 0}
            if is_pet:
                all_months_prod[mstr]["PET"] += qty
            else:
                all_months_prod[mstr]["TUBE"] += qty
            all_months_prod[mstr]["rows"] += 1

    for m, d in sorted(all_months_prod.items()):
        print(f"  {m:16s} -> TUBE: {d['TUBE']:9,} | PET: {d['PET']:9,} | Rows: {d['rows']}")

print("\n" + "="*70)
print(" 3. INSPECTING Production_Archive.xlsx FOR July 2026")
print("="*70)
arch_path = os.path.join(RECORDS_DIR, "Production_Archive.xlsx")
if os.path.exists(arch_path):
    wb = openpyxl.load_workbook(arch_path, data_only=True)
    print("Sheets in Production_Archive.xlsx:", wb.sheetnames)
    
    for sname in wb.sheetnames:
        if "July" in sname or sname in ("All Months", "Customer Breakdown"):
            ws = wb[sname]
            print(f"\n --- Sheet [{sname}] ({ws.max_row} rows) ---")
            if sname == "Customer Breakdown":
                j_tube = 0
                j_pet = 0
                for r in range(4, ws.max_row + 1):
                    c_cust = ws.cell(r, 1).value
                    c_m = ws.cell(r, 2).value
                    if c_m == "July 2026":
                        t_p = ws.cell(r, 3).value or 0
                        p_p = ws.cell(r, 5).value or 0
                        j_tube += t_p
                        j_pet += p_p
                print(f"   July 2026 in Customer Breakdown -> TUBE Prod: {j_tube:,} | PET Prod: {j_pet:,}")
            elif sname == "July 2026":
                t_sum = 0
                p_sum = 0
                for r in range(3, ws.max_row + 1):
                    pname = ws.cell(r, 4).value
                    dia   = str(ws.cell(r, 5).value or "")
                    good  = ws.cell(r, 8).value
                    if pname and good is not None:
                        qty = int(good or 0)
                        if "ml" in dia.lower(): p_sum += qty
                        else: t_sum += qty
                print(f"   July 2026 Tab Sum -> TUBE Prod: {t_sum:,} | PET Prod: {p_sum:,}")
    wb.close()

print("\n" + "="*70)
print(" 4. INSPECTING Tubex.html CUSTOMER_REPORT_DATA FOR July 2026 & All Months")
print("="*70)
html_path = os.path.join(ALPHA_DIR, "Tubex.html")
if os.path.exists(html_path):
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()
    m = re.search(r"const CUSTOMER_REPORT_DATA = (\[.*?\]);", content, re.DOTALL)
    if m:
        data = json.loads(m.group(1))
        print(f" Total Customer Report Records in Tubex.html: {len(data)}")
        
        july_recs = [r for r in data if r["month"] == "July 2026"]
        j_tube = sum(r["produced"] for r in july_recs if r["type"] == "TUBE")
        j_pet  = sum(r["produced"] for r in july_recs if r["type"] == "PET")
        print(f" Tubex.html July 2026 -> TUBE Prod: {j_tube:,} | PET Prod: {j_pet:,}")
        
        tot_tube = sum(r["produced"] for r in data if r["type"] == "TUBE")
        tot_pet  = sum(r["produced"] for r in data if r["type"] == "PET")
        print(f" Tubex.html All Months -> TUBE Prod: {tot_tube:,} | PET Prod: {tot_pet:,}")

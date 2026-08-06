"""
parse_legacy_dispatch.py
========================
Parses legacy dispatch XLS files:
  1. d:\Alpha\Tubex Records\dispatch nov to jul.xls (Tubes)
  2. d:\Alpha\Tubex Records\dispatch pet nov to jul.xls (PET)

Extracts line-by-line dispatch records and summarizes by month and customer.
"""

import os, xlrd, openpyxl, datetime

RECORDS_DIR = r"d:\Alpha\Tubex Records"
ALPHA_DIR   = r"d:\Alpha"

# Load catalog mapping
catalog_map = {} # product_name_upper -> {pid, customer, dia, pname}

aug_path = os.path.join(ALPHA_DIR, "Tubex_Aug26.xlsx")
if os.path.exists(aug_path):
    wb_cat = openpyxl.load_workbook(aug_path, data_only=True)
    if "Product_Catalog" in wb_cat.sheetnames:
        ws_cat = wb_cat["Product_Catalog"]
        for r in range(3, ws_cat.max_row + 1):
            pid   = ws_cat.cell(r, 1).value
            cust  = ws_cat.cell(r, 3).value
            pname = ws_cat.cell(r, 4).value
            dia   = ws_cat.cell(r, 5).value
            if pname:
                catalog_map[str(pname).strip().upper()] = {
                    "pid": pid,
                    "customer": str(cust).strip() if cust else "Other Customers",
                    "dia": str(dia).strip() if dia else "",
                    "pname": str(pname).strip()
                }
    wb_cat.close()

def parse_dispatch_xls(filename, default_type):
    path = os.path.join(RECORDS_DIR, filename)
    if not os.path.exists(path):
        print(f"File not found: {path}")
        return []

    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_index(0)

    records = []
    current_product = ""

    for r in range(6, ws.nrows):
        row = [ws.cell_value(r, c) for c in range(ws.ncols)]
        if not row:
            continue

        c0 = row[0]
        # Check if product header row
        if isinstance(c0, str) and c0.strip() and not c0.strip().startswith("Total") and not c0.strip().startswith("No"):
            # Could be a product header
            if row[1] == "" and row[2] == "" and row[3] == "":
                current_product = c0.strip()
                continue

        # Check if record row (has numeric No in c0 or c1)
        is_rec = False
        if isinstance(c0, float) and c0 > 0:
            is_rec = True
        elif isinstance(row[1], float) and row[1] > 0 and isinstance(row[2], float):
            is_rec = True

        if is_rec and current_product:
            date_val = row[2]
            pof      = row[3]
            client_po= row[4]
            dia      = row[5]
            ord_qty  = row[6]
            disp_qty = row[7]
            repl_qty = row[8]
            sv_no    = row[9]

            dt = None
            if isinstance(date_val, float):
                try:
                    dt = xlrd.xldate_as_datetime(date_val, wb.datemode)
                except:
                    pass

            date_str = dt.strftime("%Y-%m-%d") if dt else str(date_val)
            month_str = dt.strftime("%B %Y") if dt else "Unknown"

            disp_num = float(disp_qty) if disp_qty not in ("", None) else 0.0
            ord_num  = float(ord_qty)  if ord_qty  not in ("", None) else 0.0

            # Catalog lookup
            cat = catalog_map.get(current_product.upper(), {})
            
            from customer_normalization import normalize_customer_name, correct_customer_by_product
            alpha_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            customer = normalize_customer_name(cat.get("customer", "Other Customers"), alpha_dir)
            customer = correct_customer_by_product(customer, current_product)

            records.append({
                "type": default_type,
                "source": filename,
                "date": date_str,
                "month": month_str,
                "product": current_product,
                "customer": customer,
                "dia": str(dia).strip() or cat.get("dia", ""),
                "ord_qty": int(ord_num),
                "disp_qty": int(disp_num),
                "pof": pof,
                "client_po": client_po,
                "sv_no": sv_no
            })

    return records

tube_disp = parse_dispatch_xls("dispatch nov to jul.xls", "TUBE")
pet_disp  = parse_dispatch_xls("dispatch pet nov to jul.xls", "PET")

all_disp = tube_disp + pet_disp
print(f"Total Tube Dispatch Records: {len(tube_disp):,}")
print(f"Total PET Dispatch Records : {len(pet_disp):,}")
print(f"Total Combined Records     : {len(all_disp):,}")

# Summarize by month
monthly_tot = {}
for r in all_disp:
    m = r["month"]
    t = r["type"]
    if m not in monthly_tot:
        monthly_tot[m] = {"TUBE": 0, "PET": 0}
    monthly_tot[m][t] += r["disp_qty"]

print("\n" + "="*50)
print(" MONTHLY DISPATCH SUMMARY (Nov 2025 - Jul 2026)")
print("="*50)
for m, d in sorted(monthly_tot.items()):
    print(f" {m:16s} | TUBE: {d['TUBE']:9,} | PET: {d['PET']:9,} | TOTAL: {d['TUBE']+d['PET']:9,}")

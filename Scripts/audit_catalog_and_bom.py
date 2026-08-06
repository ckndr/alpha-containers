"""
audit_catalog_and_bom.py
========================
Audits all historical production records (Nov 2025 - Aug 2026) against:
  1. Product Catalog in Tubex_Aug26.xlsx
  2. BOM dictionary in Tubex_Aug26.xlsx

Finds:
  - Products produced that are NOT in Product Catalog
  - Products in Catalog that do NOT have BOM entries
  - Products assigned to 'Other Customers' or unmapped customer names
"""

import os, xlrd, openpyxl, datetime

RECORDS_DIR = r"d:\Alpha\Tubex Records"
ALPHA_DIR   = r"d:\Alpha"

# 1. Load Product Catalog from Tubex_Aug26.xlsx
catalog_pnames = set()
catalog_pids   = set()
catalog_by_name = {}

aug_path = r"d:\Alpha\Tubex_Aug26.xlsx"
wb_aug = openpyxl.load_workbook(aug_path, data_only=True)

if "Product_Catalog" in wb_aug.sheetnames:
    ws_cat = wb_aug["Product_Catalog"]
    for r in range(3, ws_cat.max_row + 1):
        pid   = ws_cat.cell(r, 1).value
        cust  = ws_cat.cell(r, 3).value
        pname = ws_cat.cell(r, 4).value
        dia   = ws_cat.cell(r, 5).value
        if pname:
            pname_str = str(pname).strip()
            catalog_pnames.add(pname_str.upper())
            if pid:
                catalog_pids.add(int(pid))
            catalog_by_name[pname_str.upper()] = {
                "pid": pid,
                "customer": cust,
                "pname": pname_str,
                "dia": dia
            }

# 2. Load BOM from Tubex_Aug26.xlsx
bom_pids = set()
if "BOM" in wb_aug.sheetnames:
    ws_bom = wb_aug["BOM"]
    for r in range(3, ws_bom.max_row + 1):
        pid = ws_bom.cell(r, 1).value
        if pid:
            try:
                bom_pids.add(int(pid))
            except:
                pass

wb_aug.close()

print(f"Product Catalog entries: {len(catalog_by_name)}")
print(f"BOM unique Product IDs : {len(bom_pids)}")

# 3. Read all raw production records across all files
raw_records = []

# Legacy XLS
prod_path = os.path.join(RECORDS_DIR, "production nov to jul.xls")
if os.path.exists(prod_path):
    wb = xlrd.open_workbook(prod_path)
    ws = wb.sheet_by_index(0)
    for r in range(10, ws.nrows):
        row = [ws.cell_value(r, c) for c in range(ws.ncols)]
        if not row or len(row) < 10:
            continue
        ref_no, date_val, line_no, shift, ptime, pof, pname, target, dia, good, wastage = (row + [None]*15)[:11]
        if pname and good not in (None, ""):
            dt = None
            if isinstance(date_val, float):
                try:
                    dt = xlrd.xldate_as_datetime(date_val, wb.datemode)
                except:
                    pass
            date_str = dt.strftime("%Y-%m-%d") if dt else str(date_val)
            month_str = dt.strftime("%B %Y") if dt else "Unknown"
            raw_records.append({
                "source": "production nov to jul.xls",
                "date": date_str,
                "month": month_str,
                "pof": pof,
                "product": str(pname).strip(),
                "dia": str(dia).strip() if dia else "",
                "good": int(float(good)),
                "reject": int(float(wastage or 0))
            })

# Tubex_July26.xlsx
july_path = os.path.join(RECORDS_DIR, "Tubex_July26.xlsx")
if os.path.exists(july_path):
    wb = openpyxl.load_workbook(july_path, data_only=True)
    if "Production_Log" in wb.sheetnames:
        ws = wb["Production_Log"]
        for r in range(3, ws.max_row + 1):
            row = [ws.cell(r, c).value for c in range(1, 11)]
            if row[0] and row[3]:
                date_val, machine, customer, product, dia, pid, target, good, reject, waste = (row + [None]*10)[:10]
                if good is not None:
                    dt_str = date_val.strftime("%Y-%m-%d") if isinstance(date_val, datetime.datetime) else str(date_val)
                    raw_records.append({
                        "source": "Tubex_July26.xlsx",
                        "date": dt_str,
                        "month": "July 2026",
                        "customer": str(customer).strip() if customer else "",
                        "product": str(product).strip(),
                        "dia": str(dia).strip() if dia else "",
                        "pid": pid,
                        "good": int(good or 0),
                        "reject": int(reject or 0)
                    })
    wb.close()

# Tubex_Aug26.xlsx
aug_path = os.path.join(ALPHA_DIR, "Tubex_Aug26.xlsx")
if os.path.exists(aug_path):
    wb = openpyxl.load_workbook(aug_path, data_only=True)
    if "Production_Log" in wb.sheetnames:
        ws = wb["Production_Log"]
        for r in range(3, ws.max_row + 1):
            row = [ws.cell(r, c).value for c in range(1, 11)]
            if row[0] and row[3]:
                date_val, machine, customer, product, dia, pid, target, good, reject, waste = (row + [None]*10)[:10]
                if good is not None:
                    dt_str = date_val.strftime("%Y-%m-%d") if isinstance(date_val, datetime.datetime) else str(date_val)
                    raw_records.append({
                        "source": "Tubex_Aug26.xlsx",
                        "date": dt_str,
                        "month": "August 2026",
                        "customer": str(customer).strip() if customer else "",
                        "product": str(product).strip(),
                        "dia": str(dia).strip() if dia else "",
                        "pid": pid,
                        "good": int(good or 0),
                        "reject": int(reject or 0)
                    })
    wb.close()

print(f"Total raw production entries loaded: {len(raw_records)}")

# 4. Audit against Catalog and BOM
missing_from_catalog = {}
missing_bom = {}
unmapped_customer = {}

for r in raw_records:
    pname = r["product"]
    pname_upper = pname.upper()
    good = r["good"]

    # Check Catalog
    if pname_upper not in catalog_pnames:
        if pname not in missing_from_catalog:
            missing_from_catalog[pname] = {"total_qty": 0, "months": set(), "sources": set(), "dia": r["dia"]}
        missing_from_catalog[pname]["total_qty"] += good
        missing_from_catalog[pname]["months"].add(r["month"])
        missing_from_catalog[pname]["sources"].add(r["source"])
    else:
        cat_info = catalog_by_name[pname_upper]
        pid = cat_info["pid"]
        cust = cat_info["customer"]

        # Check BOM
        if pid and int(pid) not in bom_pids:
            if pname not in missing_bom:
                missing_bom[pname] = {"pid": pid, "customer": cust, "total_qty": 0}
            missing_bom[pname]["total_qty"] += good

        # Check Customer
        if not cust or str(cust).lower() in ("other customers", "other", "unknown", "none"):
            if pname not in unmapped_customer:
                unmapped_customer[pname] = {"total_qty": 0, "dia": r["dia"]}
            unmapped_customer[pname]["total_qty"] += good

print("\n" + "="*70)
print(" 1. PRODUCTS PRODUCED BUT MISSING FROM PRODUCT CATALOG")
print("="*70)
if not missing_from_catalog:
    print("  [NONE] All produced products are in Product Catalog!")
else:
    for pname, d in sorted(missing_from_catalog.items(), key=lambda x: x[1]['total_qty'], reverse=True):
        print(f"  • Product Name : {pname}")
        print(f"    Dia/Volume   : {d['dia']}")
        print(f"    Total Produced: {d['total_qty']:,} units")
        print(f"    Months       : {', '.join(sorted(d['months']))}")
        print(f"    Sources      : {', '.join(sorted(d['sources']))}")
        print()

print("\n" + "="*70)
print(" 2. CATALOG PRODUCTS PRODUCED THAT ARE MISSING BOM ENTRIES")
print("="*70)
if not missing_bom:
    print("  [NONE] All catalog products have BOM entries!")
else:
    for pname, d in sorted(missing_bom.items(), key=lambda x: x[1]['total_qty'], reverse=True):
        print(f"  • Product Name : {pname} (PID {d['pid']})")
        print(f"    Customer     : {d['customer']}")
        print(f"    Total Produced: {d['total_qty']:,} units")
        print()

print("\n" + "="*70)
print(" 3. UNMAPPED OR MISSING CUSTOMER ASSIGNMENTS")
print("="*70)
if not unmapped_customer:
    print("  [NONE] All products have specific customer assignments!")
else:
    for pname, d in sorted(unmapped_customer.items(), key=lambda x: x[1]['total_qty'], reverse=True):
        print(f"  • Product Name : {pname}")
        print(f"    Dia/Volume   : {d['dia']}")
        print(f"    Total Produced: {d['total_qty']:,} units")
        print()

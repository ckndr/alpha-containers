"""Scan all production data for unique customers and their volumes."""
import openpyxl, datetime

sources = [
    ("Archive",   r"d:\Alpha\Tubex Records\Production_Archive.xlsx", "July 2026",   3),
    ("Active",    r"d:\Alpha\Tubex_Aug26.xlsx",                       "August 2026", 3),
]

customers = {}   # customer -> {month -> {tube, pet}}

for src_label, path, month_label, start_row in sources:
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet_name = "July 2026" if "Archive" in src_label else "Production_Log"
        if sheet_name not in wb.sheetnames:
            print(f"[SKIP] {sheet_name} not in {path}")
            wb.close()
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=start_row, values_only=True):
            if not row[0]:
                continue
            date_val, machine, customer, product, dia, pid, target, good, reject, waste = (list(row) + [None]*10)[:10]
            if not customer or not good:
                continue
            # Determine type from product/machine/dia
            prod_type = "PET" if (dia and "ml" in str(dia).lower()) else "TUBE"
            key = str(customer).strip()
            if key not in customers:
                customers[key] = {}
            if month_label not in customers[key]:
                customers[key][month_label] = {"TUBE": 0, "PET": 0}
            try:
                customers[key][month_label][prod_type] += int(good or 0)
            except:
                pass
        wb.close()
    except Exception as e:
        print(f"ERROR {path}: {e}")

print(f"\nFound {len(customers)} unique customers:\n")
for cust in sorted(customers.keys()):
    total = sum(v.get("TUBE",0)+v.get("PET",0) for v in customers[cust].values())
    months = list(customers[cust].keys())
    print(f"  {cust[:50]:52s}  total={total:>10,}  months={months}")

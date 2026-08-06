"""
parse_legacy_xls.py
===================
Parses legacy XLS files in Tubex Records:
  - production nov to jul.xls
  - dispatch nov to jul.xls
  - dispatch pet nov to jul.xls
  - Samsol_Production_and_Dispatch.xlsx

Provides clean python functions to retrieve standardized production and dispatch records
grouped by Month and Customer.
"""

import os, xlrd, datetime, openpyxl

RECORDS_DIR = r"d:\Alpha\Tubex Records"

# Product Catalog mapping from Product Name -> (Customer, Dia/Volume, Type, Prod_ID)
def load_product_catalog():
    catalog = {}
    aug_path = r"d:\Alpha\Tubex_Aug26.xlsx"
    if os.path.exists(aug_path):
        try:
            wb = openpyxl.load_workbook(aug_path, data_only=True)
            if "Product_Catalog" in wb.sheetnames:
                ws = wb["Product_Catalog"]
                for r in range(3, ws.max_row + 1):
                    pid = ws.cell(r, 1).value
                    cust = ws.cell(r, 3).value
                    pname = ws.cell(r, 4).value
                    dia = ws.cell(r, 5).value
                    if pname and cust:
                        key = str(pname).strip().upper()
                        catalog[key] = {
                            "customer": str(cust).strip(),
                            "dia": str(dia).strip() if dia else "",
                            "pid": pid or "",
                            "type": "PET" if dia and "ml" in str(dia).lower() else "TUBE"
                        }
            wb.close()
        except Exception as e:
            print(f"Catalog load warning: {e}")

    # Explicit hardcoded overrides for legacy names
    overrides = {
        "TUBES COMMON RED": {"customer": "Samsol International Private Limited", "dia": "25", "pid": 6470, "type": "TUBE"},
        "TUBES": {"customer": "Samsol International Private Limited", "dia": "25", "pid": 3726, "type": "TUBE"},
        "TUBES MEN BLUE": {"customer": "Samsol International Private Limited", "dia": "25", "pid": 6506, "type": "TUBE"},
        "TUBE COMMON PURPLE": {"customer": "Samsol International Private Limited", "dia": "25", "pid": 6532, "type": "TUBE"},
        "S 43 25MM": {"customer": "Samsol International Private Limited", "dia": "25", "pid": 3447, "type": "TUBE"},
        "SAMSOL RED 25MM": {"customer": "Samsol International Private Limited", "dia": "25", "pid": 9006, "type": "TUBE"},
        "S-43 DIA 20.5": {"customer": "Samsol International Private Limited", "dia": "20.5", "pid": 5699, "type": "TUBE"},
        "S-45 DIA 20.5": {"customer": "Samsol International Private Limited", "dia": "20.5", "pid": 5698, "type": "TUBE"},
        "S-45": {"customer": "Samsol International Private Limited", "dia": "25", "pid": 5389, "type": "TUBE"},
        "S-43 DIA 19MM": {"customer": "Samsol International Private Limited", "dia": "19", "pid": 6623, "type": "TUBE"},
        "S-45 DIA 19MM": {"customer": "Samsol International Private Limited", "dia": "19", "pid": 6624, "type": "TUBE"},
        "HELLO HAIR COLOR": {"customer": "Golden Pearl Cosmetics (PVT) LTD", "dia": "30", "pid": 6206, "type": "TUBE"},
        "GP DIA 30MM": {"customer": "Golden Pearl Cosmetics (PVT) LTD", "dia": "30", "pid": 6206, "type": "TUBE"},
        "ACTIVE PRO HAIR COLOR CREAM 60ML": {"customer": "Al-Rehman Group", "dia": "30", "pid": 6337, "type": "TUBE"},
        "SIGNATURE HAIR COLOR CREAM 60ML": {"customer": "Al-Rehman Group", "dia": "30", "pid": 6416, "type": "TUBE"},
        "VINCE NURTURAL": {"customer": "Mablay Beauty PVT LTD.", "dia": "30", "pid": 5814, "type": "TUBE"},
        "VINCE HIS ONLY BEARD & MUSTACHE COLOR CREAM": {"customer": "Mablay Beauty PVT LTD.", "dia": "20.5", "pid": 6077, "type": "TUBE"},
        "HIS ONLY HAIR COLOR CREAM 40GM": {"customer": "Mablay Beauty PVT LTD.", "dia": "25", "pid": 6228, "type": "TUBE"},
        "PET BOTTLE SMALL (120ML) YELLOW": {"customer": "Samsol International Private Limited", "dia": "120 ml", "pid": 8005, "type": "PET"},
        "PET BOTTLE LARGE (200 ML) YELLOW": {"customer": "Samsol International Private Limited", "dia": "200 ml", "pid": 8006, "type": "PET"},
        "PET BOTTLE 130ML WHITE": {"customer": "Mablay Beauty PVT LTD.", "dia": "130 ml", "pid": 8015, "type": "PET"},
        "TRANSPARENT BOTTLE 150ML": {"customer": "Alpha Labs PVT LTD", "dia": "150 ml", "pid": 8001, "type": "PET"},
        "PET BOTTLE (150ML)TRANSPARENT BODY MIST": {"customer": "Alpha Labs PVT LTD", "dia": "150 ml", "pid": 8001, "type": "PET"},
        "BT-120 ML YELLOW": {"customer": "Samsol International Private Limited", "dia": "120 ml", "pid": 8005, "type": "PET"},
        "BT-200ML MUSTARD OIL (TRANSPARENT)": {"customer": "Samsol International Private Limited", "dia": "200 ml", "pid": 8014, "type": "PET"},
    }
    for k, v in overrides.items():
        catalog[k] = v
    return catalog

def get_legacy_production_records():
    catalog = load_product_catalog()
    prod_path = os.path.join(RECORDS_DIR, "production nov to jul.xls")
    records = []

    if not os.path.exists(prod_path):
        return records

    try:
        wb = xlrd.open_workbook(prod_path)
        ws = wb.sheet_by_index(0)

        for r in range(10, ws.nrows):
            row = [ws.cell_value(r, c) for c in range(ws.ncols)]
            if not row or len(row) < 10:
                continue

            ref_no, date_val, line_no, shift, ptime, pof, pname, target, dia, good, wastage = (row + [None]*15)[:11]

            if not pname or good is None or good == "":
                continue

            # Parse date
            dt = None
            if isinstance(date_val, float):
                try:
                    dt = xlrd.xldate_as_datetime(date_val, wb.datemode)
                except:
                    pass

            if not dt:
                continue

            month_label = dt.strftime("%B %Y")
            date_str = dt.strftime("%Y-%m-%d")

            pname_str = str(pname).strip()
            cat_entry = catalog.get(pname_str.upper(), {})

            from customer_normalization import normalize_customer_name, correct_customer_by_product
            alpha_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            customer = normalize_customer_name(cat_entry.get("customer", "Other Customers"), alpha_dir)
            customer = correct_customer_by_product(customer, pname_str)
            dia_str  = str(dia).strip() if dia else cat_entry.get("dia", "")
            pid      = cat_entry.get("pid", "")
            ptype    = cat_entry.get("type", "TUBE")

            try:
                good_qty = int(float(good))
            except:
                good_qty = 0

            try:
                target_qty = int(float(target)) if target else 0
            except:
                target_qty = 0

            try:
                rej_qty = int(float(wastage)) if wastage else 0
            except:
                rej_qty = 0

            machine = f"Printing-0{int(float(line_no))}" if isinstance(line_no, (int, float)) else str(line_no or "")

            records.append({
                "date": date_str,
                "month": month_label,
                "customer": customer,
                "product": pname_str,
                "machine": machine,
                "dia": dia_str,
                "pid": pid,
                "type": ptype,
                "target": target_qty,
                "good": good_qty,
                "reject": rej_qty,
                "source": "production nov to jul.xls"
            })
    except Exception as e:
        print(f"Error reading production nov to jul.xls: {e}")

    # Parse PET production from Production report Jan-2026 till Date.xlsx
    pet_prod_path = os.path.join(RECORDS_DIR, "Production report Jan-2026 till Date.xlsx")
    if os.path.exists(pet_prod_path):
        try:
            wb = openpyxl.load_workbook(pet_prod_path, data_only=True)
            if "Production Day wise" in wb.sheetnames:
                ws = wb["Production Day wise"]
                for r in range(4, ws.max_row + 1):
                    mach = str(ws.cell(r, 3).value or "").strip()
                    if mach.startswith("PF"):
                        dt_val = ws.cell(r, 1).value
                        if not isinstance(dt_val, datetime.datetime):
                            continue
                        month_label = dt_val.strftime("%B %Y")
                        date_str = dt_val.strftime("%Y-%m-%d")
                        
                        pname = ws.cell(r, 5).value or ""
                        pname_str = str(pname).strip()
                        cat_entry = catalog.get(pname_str.upper(), {})
                        
                        cust_name = ws.cell(r, 4).value or ""
                        raw_cust = str(cust_name).strip() if cust_name else cat_entry.get("customer", "Other Customers")
                        from customer_normalization import normalize_customer_name, correct_customer_by_product
                        alpha_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                        customer = normalize_customer_name(raw_cust, alpha_dir)
                        customer = correct_customer_by_product(customer, pname_str)
                        
                        dia = ws.cell(r, 6).value or ""
                        dia_str = str(dia).strip() if dia else cat_entry.get("dia", "")
                        
                        target_qty = ws.cell(r, 14).value
                        good_qty = ws.cell(r, 12).value
                        rej_qty = ws.cell(r, 11).value
                        
                        records.append({
                            "date": date_str,
                            "month": month_label,
                            "customer": customer,
                            "product": pname_str,
                            "machine": mach,
                            "dia": dia_str,
                            "pid": cat_entry.get("pid", ""),
                            "type": "PET",
                            "target": int(target_qty) if target_qty else 0,
                            "good": int(good_qty) if good_qty else 0,
                            "reject": int(rej_qty) if rej_qty else 0,
                            "source": "Production report Jan-2026 till Date.xlsx"
                        })
            wb.close()
        except Exception as e:
            print(f"Error reading PET production report: {e}")

    return records

if __name__ == "__main__":
    recs = get_legacy_production_records()
    print(f"Loaded {len(recs)} legacy production records")
    months = sorted(set(r['month'] for r in recs))
    print(f"Months: {months}")

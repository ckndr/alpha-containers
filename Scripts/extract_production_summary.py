"""
extract_production_summary.py
==============================
Python script using `xlrd` (with `pandas` fallback) to extract and summarize
data from ERP-generated Tubes Production report:
  'D:\\Alpha\\Tubex Records\\production nov to jul.xls'

Outputs:
  1. All sheet names
  2. All column headers
  3. A structured summary grouped by Month and Customer showing total production quantities
"""

import os
import sys
import datetime
from collections import defaultdict

# Add workspace scripts directory to path for catalog & normalization imports
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.append(SCRIPT_DIR)

# File path
EXCEL_PATH = r"D:\Alpha\Tubex Records\production nov to jul.xls"
ALPHA_DIR = os.path.dirname(SCRIPT_DIR)

def load_customer_helpers():
    """Import normalization and catalog lookup functions if available."""
    try:
        from parse_legacy_xls import load_product_catalog
        from customer_normalization import normalize_customer_name, correct_customer_by_product
        catalog = load_product_catalog()
        return catalog, normalize_customer_name, correct_customer_by_product
    except Exception as e:
        print(f"Notice: Could not load master catalog module ({e}). Using built-in fallback mapping.")
        return {}, None, None

def get_customer_name(product_name, catalog, norm_func, corr_func):
    """Determine standardized customer name for a given product."""
    if not product_name:
        return "Unknown / Unmapped"
    
    pname_str = str(product_name).strip()
    pname_upper = pname_str.upper()
    
    # Check catalog first
    if catalog and pname_upper in catalog:
        raw_cust = catalog[pname_upper].get("customer")
        if raw_cust:
            cust = norm_func(raw_cust, ALPHA_DIR) if norm_func else raw_cust
            cust = corr_func(cust, pname_str) if corr_func else cust
            return cust

    # Fallback rules for known legacy products if catalog missing
    if "HELLO HAIR" in pname_upper or "GP DIA" in pname_upper:
        return "Golden Pearl Cosmetics (PVT) LTD"
    elif "SAMSOL" in pname_upper or "S-43" in pname_upper or "S-45" in pname_upper or "TUBES" in pname_upper or "S 43" in pname_upper:
        return "Samsol International Private Limited"
    elif "HIS ONLY" in pname_upper or "VINCE" in pname_upper:
        return "Mablay Beauty PVT LTD."
    elif "H.H" in pname_upper:
        return "Hola Hair"
    elif "ANVIL" in pname_upper or "V- HC" in pname_upper or "V-HC" in pname_upper:
        return "Adore"
    elif "ACTIVE PRO" in pname_upper or "SIGNATURE" in pname_upper:
        return "Al-Rehman Group"
    elif "CONTRATUBEX" in pname_upper or "PHLOGIN" in pname_upper or "PYODINE" in pname_upper:
        return "Brookes Pharma Private Limited"
    elif "DOWFEN" in pname_upper:
        return "Seatle (Private) Limited"
    elif "EAZICOLOR" in pname_upper or "HIBA S" in pname_upper:
        return "Professional Beauty Solution (Pvt) Ltd"
    elif "DARI MOOCH" in pname_upper:
        return "Dari Mooch"
    elif "M.G" in pname_upper:
        return "Mega Grey"
    elif "DTM" in pname_upper:
        return "DTM"
        
    return f"Other ({pname_str})"

def process_with_xlrd(path):
    import xlrd
    
    wb = xlrd.open_workbook(path)
    sheet_names = wb.sheet_names()
    sheet = wb.sheet_by_index(0)
    
    # 1. Sheet names
    print("=" * 80)
    print("1. SHEET NAMES")
    print("=" * 80)
    for idx, sname in enumerate(sheet_names, 1):
        print(f"  [{idx}] {sname}")
    print()
    
    # 2. Column headers
    print("=" * 80)
    print("2. COLUMN HEADERS")
    print("=" * 80)
    
    sub_headers = [sheet.cell_value(7, c) for c in range(sheet.ncols)]
    col_headers = [sheet.cell_value(9, c) for c in range(sheet.ncols)]
    
    # Build readable combined headers
    combined_headers = []
    for idx, (sub, col) in enumerate(zip(sub_headers, col_headers)):
        sub_clean = str(sub).strip()
        col_clean = str(col).strip()
        if sub_clean and sub_clean != col_clean:
            full_header = f"{sub_clean} - {col_clean}"
        else:
            full_header = col_clean
        combined_headers.append(full_header)
        print(f"  Col {idx:2d}: {full_header:<30} (Sub-header: '{sub_clean}', Main: '{col_clean}')")
    print()

    # 3. Data Extraction & Grouping
    catalog, norm_func, corr_func = load_customer_helpers()
    
    # Aggregation dictionaries
    # key: (month_sort_key, month_display, customer)
    summary_data = defaultdict(lambda: {
        'shift_qty': 0.0,
        'ot_qty': 0.0,
        'total_qty': 0.0,
        'target_qty': 0.0,
        'shift_wastage': 0.0,
        'ot_wastage': 0.0,
        'total_wastage': 0.0,
        'records': 0
    })
    
    month_summary = defaultdict(lambda: {
        'shift_qty': 0.0,
        'ot_qty': 0.0,
        'total_qty': 0.0,
        'target_qty': 0.0,
        'total_wastage': 0.0,
        'records': 0
    })
    
    grand_total = {
        'shift_qty': 0.0,
        'ot_qty': 0.0,
        'total_qty': 0.0,
        'target_qty': 0.0,
        'total_wastage': 0.0,
        'records': 0
    }

    for r in range(10, sheet.nrows):
        ref_no = sheet.cell_value(r, 0)
        # Data rows have float/numeric ref #
        if isinstance(ref_no, float):
            date_val = sheet.cell_value(r, 1)
            pname = sheet.cell_value(r, 6)
            target = float(sheet.cell_value(r, 7) or 0)
            shift_qty = float(sheet.cell_value(r, 9) or 0)
            shift_waste = float(sheet.cell_value(r, 10) or 0)
            ot_qty = float(sheet.cell_value(r, 12) or 0)
            ot_waste = float(sheet.cell_value(r, 13) or 0)
            
            tot_qty = shift_qty + ot_qty
            tot_waste = shift_waste + ot_waste
            
            # Parse Excel date
            try:
                dt = xlrd.xldate_as_datetime(date_val, wb.datemode)
                month_sort = dt.strftime("%Y-%m")
                month_disp = dt.strftime("%B %Y")
            except Exception:
                month_sort = "0000-00"
                month_disp = "Unknown Date"

            customer = get_customer_name(pname, catalog, norm_func, corr_func)
            
            key = (month_sort, month_disp, customer)
            summary_data[key]['shift_qty'] += shift_qty
            summary_data[key]['ot_qty'] += ot_qty
            summary_data[key]['total_qty'] += tot_qty
            summary_data[key]['target_qty'] += target
            summary_data[key]['shift_wastage'] += shift_waste
            summary_data[key]['ot_wastage'] += ot_waste
            summary_data[key]['total_wastage'] += tot_waste
            summary_data[key]['records'] += 1

            mkey = (month_sort, month_disp)
            month_summary[mkey]['shift_qty'] += shift_qty
            month_summary[mkey]['ot_qty'] += ot_qty
            month_summary[mkey]['total_qty'] += tot_qty
            month_summary[mkey]['target_qty'] += target
            month_summary[mkey]['total_wastage'] += tot_waste
            month_summary[mkey]['records'] += 1

            grand_total['shift_qty'] += shift_qty
            grand_total['ot_qty'] += ot_qty
            grand_total['total_qty'] += tot_qty
            grand_total['target_qty'] += target
            grand_total['total_wastage'] += tot_waste
            grand_total['records'] += 1

    # 4. Print Summary
    print("=" * 80)
    print("3. PRODUCTION SUMMARY GROUPED BY MONTH AND CUSTOMER")
    print("=" * 80)
    
    current_month_sort = None
    header_fmt = "{:<42} | {:>12} | {:>10} | {:>14} | {:>12} | {:>6}"
    row_fmt    = "{:<42} | {:>12,.0f} | {:>10,.0f} | {:>14,.0f} | {:>12,.0f} | {:>6d}"

    for (m_sort, m_disp, cust), stats in sorted(summary_data.items()):
        if m_sort != current_month_sort:
            if current_month_sort is not None:
                # Print month subtotal
                prev_mkey = [k for k in month_summary.keys() if k[0] == current_month_sort][0]
                mstats = month_summary[prev_mkey]
                print("-" * 105)
                print(row_fmt.format(f"SUBTOTAL ({prev_mkey[1]})", mstats['shift_qty'], mstats['ot_qty'], mstats['total_qty'], mstats['total_wastage'], mstats['records']))
                print()

            current_month_sort = m_sort
            print(f"--- {m_disp} ---")
            print(header_fmt.format("Customer", "Shift Qty", "OT Qty", "Total Prod Qty", "Wastage Qty", "Batch"))
            print("-" * 105)

        print(row_fmt.format(cust, stats['shift_qty'], stats['ot_qty'], stats['total_qty'], stats['total_wastage'], stats['records']))

    # Print final month subtotal
    if current_month_sort is not None:
        prev_mkey = [k for k in month_summary.keys() if k[0] == current_month_sort][0]
        mstats = month_summary[prev_mkey]
        print("-" * 105)
        print(row_fmt.format(f"SUBTOTAL ({prev_mkey[1]})", mstats['shift_qty'], mstats['ot_qty'], mstats['total_qty'], mstats['total_wastage'], mstats['records']))
        print()

    print("=" * 105)
    print(row_fmt.format("GRAND TOTAL", grand_total['shift_qty'], grand_total['ot_qty'], grand_total['total_qty'], grand_total['total_wastage'], grand_total['records']))
    print("=" * 105)

def process_with_pandas(path):
    import pandas as pd
    
    print("Using pandas fallback engine='xlrd'...")
    xl = pd.ExcelFile(path, engine='xlrd')
    print("Sheet Names:", xl.sheet_names)
    df = xl.parse(xl.sheet_names[0])
    print("Columns:", list(df.columns))
    print(df.head())

def main():
    if not os.path.exists(EXCEL_PATH):
        print(f"Error: File not found at '{EXCEL_PATH}'")
        return
        
    try:
        process_with_xlrd(EXCEL_PATH)
    except Exception as e:
        print(f"xlrd processing failed with error: {e}")
        try:
            process_with_pandas(EXCEL_PATH)
        except Exception as e2:
            print(f"pandas fallback also failed: {e2}")

if __name__ == "__main__":
    main()

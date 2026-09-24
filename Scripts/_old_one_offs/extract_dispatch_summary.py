import xlrd
from datetime import datetime
from collections import defaultdict

def extract_dispatch_summary(file_path):
    print("=" * 80)
    print("ERP DISPATCH TUBES REPORT - DATA EXTRACTION & SUMMARY")
    print("=" * 80)
    
    # 1. Open workbook with xlrd
    try:
        wb = xlrd.open_workbook(file_path)
    except Exception as e:
        print(f"xlrd failed with error: {e}")
        print("Falling back to pandas with engine='xlrd'...")
        import pandas as pd
        excel_file = pd.ExcelFile(file_path, engine='xlrd')
        sheet_names = excel_file.sheet_names
        print(f"\n1. ALL SHEET NAMES ({len(sheet_names)}):")
        for s in sheet_names:
            print(f"   - {s}")
        df = pd.read_excel(file_path, sheet_name=sheet_names[0], engine='xlrd')
        print(f"\n2. COLUMN HEADERS:")
        headers = df.iloc[5].tolist()
        for idx, h in enumerate(headers):
            print(f"   Column {idx:2d}: {h}")
        return

    # 1. Sheet names
    sheet_names = wb.sheet_names()
    print(f"\n1. ALL SHEET NAMES ({len(sheet_names)}):")
    for s_name in sheet_names:
        print(f"   - {s_name}")

    sheet = wb.sheet_by_name(sheet_names[0])

    # 2. Column headers (Found at Row index 5 in ERP layout)
    header_row_idx = 5
    headers = sheet.row_values(header_row_idx)
    print(f"\n2. ALL COLUMN HEADERS (Row {header_row_idx + 1}):")
    for idx, h in enumerate(headers):
        print(f"   Col {idx:2d}: {str(h).strip()}")

    # 3. Parse detail records
    # Group data by (YearMonth, MonthName, Customer)
    summary_data = defaultdict(lambda: {'total_qty': 0.0, 'count': 0})
    monthly_totals = defaultdict(lambda: {'total_qty': 0.0, 'count': 0})
    customer_totals = defaultdict(lambda: {'total_qty': 0.0, 'count': 0})
    grand_total_qty = 0.0
    total_records = 0

    for r in range(sheet.nrows):
        row = sheet.row_values(r)
        c0 = row[0]

        # Detail rows start with a float/int 'No.' in column 0 (e.g. 1.0, 2.0, ..., 171.0)
        if isinstance(c0, (int, float)) and 0 < c0 < 1000:
            date_val = row[2]
            date_obj = None
            if isinstance(date_val, (int, float)) and date_val > 0:
                dt_tuple = xlrd.xldate_as_tuple(date_val, wb.datemode)
                date_obj = datetime(*dt_tuple[:6])

            # Dispatch Quantity (Column Index 7)
            disp_qty_raw = row[7]
            try:
                disp_qty = float(disp_qty_raw) if disp_qty_raw != '' else 0.0
            except (ValueError, TypeError):
                disp_qty = 0.0

            # Party Name (Column Index 12)
            customer = str(row[12]).strip()
            if not customer:
                customer = "UNKNOWN / UNASSIGNED"

            year_month = date_obj.strftime('%Y-%m') if date_obj else 'Unknown'
            month_name = date_obj.strftime('%b %Y') if date_obj else 'Unknown'

            # Accumulate totals
            summary_data[(year_month, month_name, customer)]['total_qty'] += disp_qty
            summary_data[(year_month, month_name, customer)]['count'] += 1

            monthly_totals[(year_month, month_name)]['total_qty'] += disp_qty
            monthly_totals[(year_month, month_name)]['count'] += 1

            customer_totals[customer]['total_qty'] += disp_qty
            customer_totals[customer]['count'] += 1

            grand_total_qty += disp_qty
            total_records += 1

    # 3. Print Structured Summary
    print(f"\n3. SUMMARY OF DISPATCH DATA GROUPED BY MONTH AND CUSTOMER")
    print("-" * 80)
    print(f"{'Month':<12} | {'Customer Name':<45} | {'Txn Count':<9} | {'Total Dispatch Qty':>18}")
    print("-" * 80)

    # Sort months chronologically, then customer alphabetically
    sorted_months = sorted(set(ym for ym, mn, cust in summary_data.keys()))

    for ym in sorted_months:
        # Get all entries for this month
        month_entries = [(m_name, cust, data) for (y_m, m_name, cust), data in summary_data.items() if y_m == ym]
        month_entries.sort(key=lambda x: x[1])  # Sort by customer name
        
        m_name = month_entries[0][0]
        m_tot = monthly_totals[(ym, m_name)]
        
        for idx, (mn, cust, data) in enumerate(month_entries):
            print(f"{mn:<12} | {cust:<45} | {data['count']:<9d} | {data['total_qty']:>18,.0f}")
        
        # Subtotal per month
        print(f"  {'-> Subtotal ' + mn:<59} | {m_tot['count']:<9d} | {m_tot['total_qty']:>18,.0f}")
        print("-" * 80)

    print(f"{'GRAND TOTAL':<59} | {total_records:<9d} | {grand_total_qty:>18,.0f}")
    print("=" * 80)

    # Print Top Customers Summary
    print("\nSUMMARY BY CUSTOMER (OVERALL)")
    print("-" * 75)
    print(f"{'Customer Name':<45} | {'Txn Count':<9} | {'Total Dispatch Qty':>15}")
    print("-" * 75)
    sorted_customers = sorted(customer_totals.items(), key=lambda x: x[1]['total_qty'], reverse=True)
    for cust, data in sorted_customers:
        print(f"{cust:<45} | {data['count']:<9d} | {data['total_qty']:>15,.0f}")
    print("-" * 75)

if __name__ == "__main__":
    file_path = r'D:\Alpha\Tubex Records\dispatch nov to jul.xls'
    extract_dispatch_summary(file_path)

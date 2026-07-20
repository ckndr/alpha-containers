import pandas as pd
import sys
import os

sys.path.append("D:/Alpha/Scripts")
import update_production
import update_dispatch

catalog_path = "D:/Alpha/Tubex_July26.xlsx"
prod_path = "D:/Alpha/Tubex Records/production nov to jul.xls"
disp_path = "D:/Alpha/Tubex Records/dispatch nov to jul.xls"

# 1. Load Catalog
xls_cat = pd.ExcelFile(catalog_path)
df_cat = pd.read_excel(xls_cat, sheet_name="Product_Catalog", skiprows=1)

catalog_map = {}
for idx, row in df_cat.iterrows():
    name = str(row['Product Name']).strip().upper()
    pid = row['Product ID']
    cust = str(row['Customer']).strip()
    if pd.notna(pid):
        pid = int(pid)
    catalog_map[name] = {'pid': pid, 'customer': cust, 'name': row['Product Name']}

# 2. Parse Production File
df_p = pd.read_excel(prod_path, header=None)
prod_records = []

for idx, row in df_p.iterrows():
    row_val = row.dropna().tolist()
    if not row_val:
        continue
    dt = row[1]
    if pd.isna(dt) or not (isinstance(dt, pd.Timestamp) or hasattr(dt, 'date')):
        continue
    
    ref_num = row[0]
    date_val = pd.Timestamp(dt)
    pof_num = row[5]
    name_raw = str(row[6]).strip()
    dia_raw = str(row[8]).strip()
    
    def get_float(val):
        if pd.isna(val): return 0.0
        try: return float(val)
        except: return 0.0
        
    normal_good = get_float(row[9])
    ot_good = get_float(row[12])
    
    total_good = int(normal_good + ot_good)
    
    # Resolve product
    name_key = name_raw.lower().strip()
    match = update_production.ALIASES.get((name_key, dia_raw))
    if match is None:
        match = update_production.ALIASES.get((name_raw.lower().rstrip(), dia_raw))
        
    resolved_name = None
    resolved_pid = None
    resolved_cust = None
    
    if match:
        resolved_name, resolved_pid = match
        if resolved_pid in update_production.PID_TO_CUSTOMER:
            resolved_cust = update_production.PID_TO_CUSTOMER[resolved_pid]
        elif resolved_name.upper() in catalog_map:
            resolved_cust = catalog_map[resolved_name.upper()]['customer']
    else:
        name_upper = name_raw.upper()
        if name_upper in catalog_map:
            resolved_name = catalog_map[name_upper]['name']
            resolved_pid = catalog_map[name_upper]['pid']
            resolved_cust = catalog_map[name_upper]['customer']
            if resolved_pid in update_production.PID_TO_CUSTOMER:
                resolved_cust = update_production.PID_TO_CUSTOMER[resolved_pid]
                
    if resolved_name:
        cat_info = catalog_map.get(resolved_name.upper())
        cat_cust = cat_info['customer'] if cat_info else resolved_cust
        
        is_samsol = cat_cust and 'samsol' in str(cat_cust).lower()
        
        if is_samsol:
            prod_records.append({
                'Month': date_val.strftime("%Y-%m"),
                'Product Name': resolved_name,
                'Customer': cat_cust,
                'Produced Qty': total_good
            })

# 3. Parse Dispatch File
df_d = pd.read_excel(disp_path, header=None)
disp_records = []
current_prod = None
SKIP_PREFIXES = ('dispatch report', 'month :', 'no.')
SKIP_EXACT = {'end of file'}

for idx, row in df_d.iterrows():
    col0 = row[0]
    if isinstance(col0, str) and col0.strip():
        c0 = col0.strip()
        c0_lower = c0.lower()
        if c0_lower in SKIP_EXACT or any(c0_lower.startswith(p) for p in SKIP_PREFIXES) or 'grand total' in c0_lower:
            continue
        if pd.isna(row[1]):
            current_prod = c0
            continue
            
    if current_prod:
        try:
            int(float(col0)) # check if it's a record row
            date_val = pd.Timestamp(row[2])
            qty = get_float(row[7]) # Disp. Qty
            
            # Resolve product
            resolved_name = update_dispatch.NAME_FIXES.get(current_prod, current_prod)
            resolved_pid = None
            resolved_cust = None
            
            name_upper = resolved_name.strip().upper()
            if name_upper in catalog_map:
                catalog_entry = catalog_map[name_upper]
                resolved_name = catalog_entry['name']
                resolved_pid = catalog_entry['pid']
                resolved_cust = catalog_entry['customer']
                if resolved_pid in update_production.PID_TO_CUSTOMER:
                    resolved_cust = update_production.PID_TO_CUSTOMER[resolved_pid]
            else:
                resolved_name = None
                
            if resolved_name:
                cat_info = catalog_map.get(resolved_name.upper())
                cat_cust = cat_info['customer'] if cat_info else resolved_cust
                
                is_samsol = cat_cust and 'samsol' in str(cat_cust).lower()
                
                if is_samsol:
                    disp_records.append({
                        'Month': date_val.strftime("%Y-%m"),
                        'Product Name': resolved_name,
                        'Customer': cat_cust,
                        'Dispatched Qty': int(qty)
                    })
        except (ValueError, TypeError):
            pass

df_p = pd.DataFrame(prod_records)
df_d = pd.DataFrame(disp_records)

# Group by Product and Month
df_p_grouped = df_p.groupby(['Month', 'Product Name', 'Customer'])['Produced Qty'].sum().reset_index()
df_d_grouped = df_d.groupby(['Month', 'Product Name', 'Customer'])['Dispatched Qty'].sum().reset_index()

# Merge monthly
df_monthly = pd.merge(df_p_grouped, df_d_grouped, on=['Month', 'Product Name', 'Customer'], how='outer').fillna(0)
df_monthly['Produced Qty'] = df_monthly['Produced Qty'].astype(int)
df_monthly['Dispatched Qty'] = df_monthly['Dispatched Qty'].astype(int)

# Sort
df_monthly = df_monthly.sort_values(['Month', 'Product Name'])

# Months mapping for nice display
months_sorted = sorted(df_monthly['Month'].unique())

print("=== MONTHLY BREAKDOWN ===")
for m in months_sorted:
    df_m = df_monthly[df_monthly['Month'] == m]
    print(f"\n--- Month: {m} ---")
    tot_p = 0
    tot_d = 0
    for idx, row in df_m.iterrows():
        p_name = row['Product Name']
        cust = row['Customer']
        prod = row['Produced Qty']
        disp = row['Dispatched Qty']
        tot_p += prod
        tot_d += disp
        print(f"  {p_name:<40} ({cust}): Produced={prod:7,d} | Dispatched={disp:7,d}")
    print(f"  {'Total':<40}: Produced={tot_p:7,d} | Dispatched={tot_d:7,d}")

# Overall breakdown
print("\n=== TOTAL BREAKDOWN (NOV TO JUL) ===")
df_p_tot = df_p.groupby(['Product Name', 'Customer'])['Produced Qty'].sum().reset_index()
df_d_tot = df_d.groupby(['Product Name', 'Customer'])['Dispatched Qty'].sum().reset_index()
df_overall = pd.merge(df_p_tot, df_d_tot, on=['Product Name', 'Customer'], how='outer').fillna(0)
df_overall['Produced Qty'] = df_overall['Produced Qty'].astype(int)
df_overall['Dispatched Qty'] = df_overall['Dispatched Qty'].astype(int)
df_overall = df_overall.sort_values('Product Name')

tot_p = 0
tot_d = 0
for idx, row in df_overall.iterrows():
    p_name = row['Product Name']
    cust = row['Customer']
    prod = row['Produced Qty']
    disp = row['Dispatched Qty']
    tot_p += prod
    tot_d += disp
    print(f"  {p_name:<40} ({cust}): Produced={prod:9,d} | Dispatched={disp:9,d}")
print(f"  {'Total':<40}: Produced={tot_p:9,d} | Dispatched={tot_d:9,d}")

import pandas as pd
import sys
import os

sys.path.append("D:/Alpha/Scripts")
import update_production
import update_dispatch

catalog_path = "D:/Alpha/Tubex_July26.xlsx"
prod_path = "D:/Alpha/Tubex Records/production nov to jul.xls"
disp_path = "D:/Alpha/Tubex Records/dispatch nov to jul.xls"

# 1. Load Catalog and determine Samsol Products
df_cat = pd.read_excel(catalog_path, sheet_name="Product_Catalog", skiprows=1)
catalog_map = {}
for idx, row in df_cat.iterrows():
    name = str(row['Product Name']).strip().upper()
    pid = row['Product ID']
    cust = str(row['Customer']).strip()
    if pd.notna(pid):
        pid = int(pid)
    catalog_map[name] = {'pid': pid, 'customer': cust, 'name': row['Product Name']}

# List of Samsol Products from catalog
samsol_pids = set()
samsol_catalog_names = set()
for name, info in catalog_map.items():
    cust = info['customer'].lower()
    if 'samsol' in cust:
        samsol_catalog_names.add(name)
        if info['pid'] is not None:
            samsol_pids.add(info['pid'])

print(f"Samsol Catalog Names: {samsol_catalog_names}")
print(f"Samsol PIDs: {samsol_pids}")

# 2. Parse Production
df_p = pd.read_excel(prod_path, header=None)
prod_records = []
unidentified_prod = set()

for idx in range(10, len(df_p)):
    row = df_p.iloc[idx]
    dt = row[1]
    if pd.isna(dt) or not (isinstance(dt, pd.Timestamp) or hasattr(dt, 'date')):
        continue
    
    # Extract data
    ref_num = row[0]
    date_val = pd.Timestamp(dt)
    pof_num = row[5]
    name_raw = str(row[6]).strip()
    dia_raw = str(row[8]).strip()
    
    # Production values:
    # Col 9: Normal shift good qty
    # Col 12: OT shift good qty
    # Col 10: Normal shift wastage
    # Col 13: OT shift wastage
    
    def get_float(val):
        if pd.isna(val): return 0.0
        try: return float(val)
        except: return 0.0
        
    normal_good = get_float(row[9])
    ot_good = get_float(row[12])
    normal_wastage = get_float(row[10])
    ot_wastage = get_float(row[13])
    
    total_good = normal_good + ot_good
    total_wastage = normal_wastage + ot_wastage
    
    # Try to resolve via ALIASES or Catalog
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
        is_samsol = resolved_cust and 'samsol' in resolved_cust.lower()
        prod_records.append({
            'date': date_val,
            'month': date_val.strftime("%Y-%m"),
            'product_name': resolved_name,
            'dia': dia_raw,
            'pof': pof_num,
            'ref': ref_num,
            'qty': total_good,
            'wastage': total_wastage,
            'is_samsol': is_samsol,
            'source_name': name_raw,
            'resolved_cust': resolved_cust
        })
    else:
        unidentified_prod.add((name_raw, dia_raw, "Production"))
        prod_records.append({
            'date': date_val,
            'month': date_val.strftime("%Y-%m"),
            'product_name': name_raw,
            'dia': dia_raw,
            'pof': pof_num,
            'ref': ref_num,
            'qty': total_good,
            'wastage': total_wastage,
            'is_samsol': False,
            'source_name': name_raw,
            'resolved_cust': "Unidentified"
        })

print(f"\nProduction parsed. Total rows: {len(prod_records)}")

# 3. Parse Dispatch
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
            repl_qty = get_float(row[8]) # Repl. Qty
            party = str(row[12]).strip()
            pof = row[3]
            ref_num = row[1] # Pk Id / Dly Id
            dia_raw = str(row[5]).strip()
            
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
                is_samsol = resolved_cust and 'samsol' in resolved_cust.lower()
                disp_records.append({
                    'date': date_val,
                    'month': date_val.strftime("%Y-%m"),
                    'product_name': resolved_name,
                    'dia': dia_raw,
                    'pof': pof,
                    'ref': ref_num,
                    'qty': qty,
                    'repl_qty': repl_qty,
                    'is_samsol': is_samsol,
                    'source_name': current_prod,
                    'party': party,
                    'resolved_cust': resolved_cust
                })
            else:
                unidentified_prod.add((current_prod, "", "Dispatch"))
                disp_records.append({
                    'date': date_val,
                    'month': date_val.strftime("%Y-%m"),
                    'product_name': current_prod,
                    'dia': dia_raw,
                    'pof': pof,
                    'ref': ref_num,
                    'qty': qty,
                    'repl_qty': repl_qty,
                    'is_samsol': False,
                    'source_name': current_prod,
                    'party': party,
                    'resolved_cust': "Unidentified"
                })
        except (ValueError, TypeError) as e:
            pass

print(f"Dispatch parsed. Total rows: {len(disp_records)}")
print("\nUnidentified Products List:")
for name, dia, src in sorted(unidentified_prod):
    print(f" - [{src}] '{name}' (dia: {dia})")

df_p_all = pd.DataFrame(prod_records)
df_d_all = pd.DataFrame(disp_records)

# Filter to Samsol products
df_p_samsol = df_p_all[df_p_all['is_samsol']]
df_d_samsol = df_d_all[df_d_all['is_samsol']]

print(f"\nSamsol Production Rows: {len(df_p_samsol)}")
print(f"Samsol Dispatch Rows: {len(df_d_samsol)}")

# Let's print monthly summary of Samsol
print("\n=== Samsol Monthly Summary ===")
prod_monthly = df_p_samsol.groupby('month')['qty'].sum().reset_index()
disp_monthly = df_d_samsol.groupby('month')['qty'].sum().reset_index()

summary = pd.merge(prod_monthly, disp_monthly, on='month', how='outer', suffixes=('_produced', '_dispatched')).fillna(0)
summary = summary.sort_values('month')
print(summary)
print("\nTotals:")
print(f"Total Produced: {summary['qty_produced'].sum()}")
print(f"Total Dispatched: {summary['qty_dispatched'].sum()}")

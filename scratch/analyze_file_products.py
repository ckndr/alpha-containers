import pandas as pd
import sys
import os

# Load mappings from update_production.py
sys.path.append("D:/Alpha/Scripts")
import update_production
import update_dispatch

catalog_path = "D:/Alpha/Tubex_July26.xlsx"
prod_path = "D:/Alpha/Tubex Records/production nov to jul.xls"
disp_path = "D:/Alpha/Tubex Records/dispatch nov to jul.xls"

# Load catalog
xls_cat = pd.ExcelFile(catalog_path)
df_cat = pd.read_excel(xls_cat, sheet_name="Product_Catalog", skiprows=1)

# Catalog product names (case-insensitive set)
catalog_names = set(df_cat['Product Name'].dropna().astype(str).str.strip().str.upper())
catalog_pids = set(df_cat['Product ID'].dropna().astype(int))

# Build catalog name to row map for lookup
catalog_map = {}
for idx, row in df_cat.iterrows():
    name = str(row['Product Name']).strip().upper()
    pid = row['Product ID']
    cust = str(row['Customer']).strip()
    if pd.notna(pid):
        pid = int(pid)
    catalog_map[name] = {'pid': pid, 'customer': cust, 'name': row['Product Name']}

print(f"Loaded {len(catalog_map)} unique products from Product Catalog.")

# 1. Read production products
df_prod = pd.read_excel(prod_path, header=None)
# Prod has headers in row 9, data from row 10
# Col 6 is Product Name, Col 8 is Dia (which we need for ALIASES)
prod_entries = []
for idx in range(10, len(df_prod)):
    row = df_prod.iloc[idx]
    # Check if this row has a valid date in Col 1
    dt = row[1]
    if pd.isna(dt) or not (isinstance(dt, pd.Timestamp) or hasattr(dt, 'date')):
        continue
    name_raw = str(row[6]).strip()
    dia_raw = str(row[8]).strip()
    prod_entries.append((name_raw, dia_raw))

prod_unique = set(prod_entries)
print(f"Unique product-dia pairs in Production: {len(prod_unique)}")

# 2. Read dispatch products
df_disp = pd.read_excel(disp_path, header=None)
disp_products = set()
current_prod = None
SKIP_PREFIXES = ('dispatch report', 'month :', 'no.')
SKIP_EXACT = {'end of file'}

for idx, row in df_disp.iterrows():
    col0 = row[0]
    if isinstance(col0, str) and col0.strip():
        c0 = col0.strip()
        c0_lower = c0.lower()
        if c0_lower in SKIP_EXACT:
            continue
        if any(c0_lower.startswith(p) for p in SKIP_PREFIXES):
            continue
        if 'grand total' in c0_lower:
            continue
        if pd.isna(row[1]):
            current_prod = c0
            disp_products.add(current_prod)

print(f"Unique products in Dispatch: {len(disp_products)}")

# Let's perform resolution
unidentified_prod = []
samsol_prod = []
other_prod = []

# Production resolution
print("\n--- Resolving Production Products ---")
for name_raw, dia_raw in sorted(prod_unique):
    # Try to resolve via ALIASES
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
        # Check if matches catalog directly
        name_upper = name_raw.strip().upper()
        if name_upper in catalog_map:
            resolved_name = catalog_map[name_upper]['name']
            resolved_pid = catalog_map[name_upper]['pid']
            resolved_cust = catalog_map[name_upper]['customer']
            if resolved_pid in update_production.PID_TO_CUSTOMER:
                resolved_cust = update_production.PID_TO_CUSTOMER[resolved_pid]
    
    # If resolved
    if resolved_name:
        is_samsol = resolved_cust and 'samsol' in resolved_cust.lower()
        if is_samsol:
            samsol_prod.append(('production', name_raw, dia_raw, resolved_name, resolved_pid, resolved_cust))
        else:
            other_prod.append(('production', name_raw, dia_raw, resolved_name, resolved_pid, resolved_cust))
    else:
        # Unidentified
        unidentified_prod.append(('production', name_raw, dia_raw, None, None, None))

# Dispatch resolution
print("\n--- Resolving Dispatch Products ---")
for name_raw in sorted(disp_products):
    # Try to resolve via NAME_FIXES
    resolved_name = update_dispatch.NAME_FIXES.get(name_raw, name_raw)
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
        resolved_name = None # could not map to catalog
        
    if resolved_name:
        is_samsol = resolved_cust and 'samsol' in resolved_cust.lower()
        if is_samsol:
            samsol_prod.append(('dispatch', name_raw, '', resolved_name, resolved_pid, resolved_cust))
        else:
            other_prod.append(('dispatch', name_raw, '', resolved_name, resolved_pid, resolved_cust))
    else:
        unidentified_prod.append(('dispatch', name_raw, '', None, None, None))

print(f"\nSamsol Products found ({len(samsol_prod)}):")
for src, raw_n, dia, res_n, pid, cust in samsol_prod:
    print(f" - [{src}] '{raw_n}' (dia: {dia}) -> '{res_n}' (PID: {pid})")

print(f"\nUnidentified Products (not in catalog) ({len(unidentified_prod)}):")
for src, raw_n, dia, _, _, _ in unidentified_prod:
    print(f" - [{src}] '{raw_n}' (dia: {dia})")

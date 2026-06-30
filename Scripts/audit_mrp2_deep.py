import pandas as pd

file_path = r'd:\Alpha\Tubex_v10_30.xlsx'

# Read full inventory including WIP
df_inv = pd.read_excel(file_path, sheet_name='Inventory', header=None)

item_id_col = None
store_bal_col = None
wip_col = None

inv_full = {}
for _, row in df_inv.iterrows():
    for i, val in enumerate(row):
        if str(val).strip() == 'Item ID':
            item_id_col = i
        elif str(val).strip() == 'Store Balance':
            store_bal_col = i
        elif str(val).strip() == 'Work In Process':
            wip_col = i
    if item_id_col is not None and store_bal_col is not None:
        try:
            item_id = str(int(row[item_id_col]))
            store_bal = row[store_bal_col]
            wip = row[wip_col] if wip_col is not None else 0
            if pd.isna(wip):
                wip = 0
            if pd.isna(store_bal):
                store_bal = 0
            inv_full[item_id] = {'store': float(store_bal), 'wip': float(wip), 'total': float(store_bal) + float(wip)}
        except:
            pass

# Items used in MRP 2
target_items = ['6', '406', '3595', '68', '578', '3598', '6935', '4155', '185', '186', '194']
print("Item ID | Store Balance | WIP | Store + WIP")
print("-" * 55)
for item_id in target_items:
    if item_id in inv_full:
        d = inv_full[item_id]
        print(f"{item_id:>7} | {d['store']:>13.2f} | {d['wip']:>5.1f} | {d['total']:>11.2f}")
    else:
        print(f"{item_id:>7} | NOT FOUND")

# Now check MRP sheet values
print("\n\nExisting MRP sheet material rows (for comparison):")
df_mrp_full = pd.read_excel(file_path, sheet_name='MRP', skiprows=1)
for _, row in df_mrp_full.iterrows():
    try:
        name = str(row.iloc[2])  # Item Name
        if name != 'nan':
            print(f"  {row.iloc[0]:>6} | {row.iloc[1]:>12} | {name:>35} | {row.iloc[3]:>5} | Req: {row.iloc[4]:>12} | Store: {row.iloc[5]:>10} | Net: {row.iloc[6]:>12} | {row.iloc[7]}")
    except:
        pass

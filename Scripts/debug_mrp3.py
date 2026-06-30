import pandas as pd

file_path = r'd:\Alpha\Tubex_v10_30.xlsx'
df_inv = pd.read_excel(file_path, sheet_name='Inventory', header=None)

inv_map = {}
item_id_col = None
store_bal_col = None

for _, row in df_inv.iterrows():
    # Detect columns
    for i, val in enumerate(row):
        if str(val).strip() == 'Item ID':
            item_id_col = i
        elif str(val).strip() == 'Store Balance':
            store_bal_col = i
            
    if item_id_col is not None and store_bal_col is not None:
        try:
            item_id = str(int(row[item_id_col]))
            store_bal = row[store_bal_col]
            inv_map[item_id] = store_bal
        except:
            pass

print("Inventory Map size:", len(inv_map))
print("Inventory Map:", inv_map)

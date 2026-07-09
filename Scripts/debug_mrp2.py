import pandas as pd

file_path = r'd:\Alpha\Tubex_July26.xlsx'
df_inv = pd.read_excel(file_path, sheet_name='Inventory', skiprows=11)
inv_map = {}
for _, row in df_inv.iterrows():
    try:
        if pd.notna(row['Item ID']):
            inv_map[str(int(row['Item ID']))] = row['Store Balance']
    except:
        pass

print("Inventory Map:", inv_map)

df_bom = pd.read_excel(file_path, sheet_name='BOM', skiprows=1)
products = ['S-45', 'S 43 25MM', 'VINCE NURTURAL', 'HELLO HAIR COLOR']
bom_rows = df_bom[df_bom['Product Name'].isin(products)]

for _, row in bom_rows.head(5).iterrows():
    print("BOM Item ID:", str(int(row['Item ID'])) if pd.notna(row['Item ID']) else 'Unknown')

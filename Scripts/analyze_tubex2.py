import pandas as pd

file_path = r'd:\Alpha\Tubex_July26.xlsx'
df_mrp = pd.read_excel(file_path, sheet_name='MRP')
print("MRP Rows 8-30:")
print(df_mrp.iloc[8:30].to_string())

print("\nInventory Sheet:")
df_inv = pd.read_excel(file_path, sheet_name='Inventory')
print(df_inv.head(20).to_string())

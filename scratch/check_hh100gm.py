import pandas as pd

catalog_path = "D:/Alpha/Tubex_July26.xlsx"
df_cat = pd.read_excel(catalog_path, sheet_name="Product_Catalog", skiprows=1)

row_pid = df_cat[df_cat['Product ID'] == 6515]
print("=== PID 6515 in Product Catalog ===")
print(row_pid[['Product ID', 'Customer', 'Product Name']])

row_name = df_cat[df_cat['Product Name'].astype(str).str.contains('H.H 100GM|Hola', case=False)]
print("\n=== Products with Name containing H.H 100GM or Hola ===")
print(row_name[['Product ID', 'Customer', 'Product Name']])

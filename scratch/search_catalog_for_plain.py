import pandas as pd

catalog_path = "D:/Alpha/Tubex_July26.xlsx"
df_cat = pd.read_excel(catalog_path, sheet_name="Product_Catalog", skiprows=1)

print("=== Searching Product Catalog for PLAIN ===")
matches_cat = df_cat[df_cat.astype(str).apply(lambda x: x.str.contains('plain|dia 16|dia 19', case=False)).any(axis=1)]
print(matches_cat[['Product ID', 'Customer', 'Product Name']])

df_bom = pd.read_excel(catalog_path, sheet_name="BOM", skiprows=1)
print("\n=== Searching BOM for PLAIN ===")
matches_bom = df_bom[df_bom.astype(str).apply(lambda x: x.str.contains('plain|dia 16|dia 19', case=False)).any(axis=1)]
print(matches_bom[['Product ID', 'Customer', 'Product Name']].drop_duplicates())

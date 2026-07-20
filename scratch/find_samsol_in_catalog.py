import pandas as pd

catalog_path = "D:/Alpha/Tubex_July26.xlsx"

# 1. Product_Catalog
df_cat = pd.read_excel(catalog_path, sheet_name="Product_Catalog", skiprows=1)
samsol_cat = df_cat[df_cat['Customer'].astype(str).str.lower().str.contains('samsol')]
print("=== Samsol Products in Product_Catalog ===")
for idx, row in samsol_cat.iterrows():
    print(f"PID: {row['Product ID']} | Name: {row['Product Name']} | Customer: {row['Customer']}")

# 2. BOM
df_bom = pd.read_excel(catalog_path, sheet_name="BOM", skiprows=1)
samsol_bom = df_bom[df_bom['Customer'].astype(str).str.lower().str.contains('samsol')]
print("\n=== Samsol Products in BOM ===")
samsol_bom_unique = samsol_bom[['Product ID', 'Product Name', 'Customer']].drop_duplicates()
for idx, row in samsol_bom_unique.iterrows():
    print(f"PID: {row['Product ID']} | Name: {row['Product Name']} | Customer: {row['Customer']}")

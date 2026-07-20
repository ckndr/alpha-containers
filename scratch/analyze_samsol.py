import pandas as pd
import os

catalog_path = "D:/Alpha/Tubex_July26.xlsx"
prod_path = "D:/Alpha/Tubex Records/production nov to jul.xls"
disp_path = "D:/Alpha/Tubex Records/dispatch nov to jul.xls"

# 1. Inspect Product Catalog for Samsol
df_cat = pd.read_excel(catalog_path, sheet_name="Product_Catalog", skiprows=1)
print("Product Catalog columns:", df_cat.columns.tolist())

# Get all unique customers
customers = df_cat['Customer'].dropna().unique()
print("\nUnique customers in Catalog:")
for c in customers:
    print(" -", c)

samsol_customers = [c for c in customers if 'samsol' in str(c).lower()]
print("\nSamsol customers identified:", samsol_customers)

samsol_catalog_products = df_cat[df_cat['Customer'].isin(samsol_customers)]
print(f"\nSamsol Catalog Products Count: {len(samsol_catalog_products)}")
print(samsol_catalog_products[['Product ID', 'BOM ID', 'Customer', 'Product Name']])

# 2. Inspect production and dispatch files
print("\n=================== Production File ===================")
xls_prod = pd.ExcelFile(prod_path)
print("Production sheet names:", xls_prod.sheet_names)
for sheet in xls_prod.sheet_names:
    df_p = pd.read_excel(prod_path, sheet_name=sheet, nrows=5)
    print(f"\nSheet {sheet} columns:", df_p.columns.tolist())
    print("Row samples:")
    print(df_p.head(2))

print("\n=================== Dispatch File ===================")
xls_disp = pd.ExcelFile(disp_path)
print("Dispatch sheet names:", xls_disp.sheet_names)
for sheet in xls_disp.sheet_names:
    df_d = pd.read_excel(disp_path, sheet_name=sheet, nrows=5)
    print(f"\nSheet {sheet} columns:", df_d.columns.tolist())
    print("Row samples:")
    print(df_d.head(2))

import pandas as pd

file_path = r'd:\Alpha\Tubex_July26.xlsx'
df_catalog = pd.read_excel(file_path, sheet_name='Product_Catalog', skiprows=1)
df_mrp = pd.read_excel(file_path, sheet_name='MRP', skiprows=1)

# Find the products in MRP to get their remaining balance
products = ['S-45', 'S 43 25MM', 'VINCE NURTURAL', 'HELLO HAIR COLOR']
mrp_rows = df_mrp[df_mrp['Product Name'].isin(products)]
print("Target Products from MRP:")
print(mrp_rows[['Product Name', 'Remaining Balance']])

balances = dict(zip(mrp_rows['Product Name'], mrp_rows['Remaining Balance']))

# Now get the BOM from Catalog
catalog_rows = df_catalog[df_catalog['Product Name'].isin(products)]
print("\nBOM for Target Products:")
print(catalog_rows[['Product Name', '# of Pieces', 'Slug (kg)', 'Base Coat (kg)', 'Lacquer (kg)', 'Latex (kg)', 'Zinc (kg)', 'Caps (PCS)', 'Cartons (PCS)', 'Ink Used']])

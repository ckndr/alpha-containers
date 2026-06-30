import pandas as pd

file_path = r'd:\Alpha\Tubex_v10_30.xlsx'
df_mrp = pd.read_excel(file_path, sheet_name='MRP', skiprows=1) # First row is likely a title
print("MRP Columns:", df_mrp.columns.tolist())
print(df_mrp.head(10).to_string())

# Let's see if we can find these specific tubes in MRP or Product_Catalog
df_catalog = pd.read_excel(file_path, sheet_name='Product_Catalog', skiprows=1)
print("\nCatalog Columns:", df_catalog.columns.tolist())

# search for S-45, S 43 25MM, VINCE NURTURAL, HELLO HAIR COLOR in catalog
search_terms = ['S-45', 'S 43 25MM', 'VINCE', 'HELLO HAIR COLOR']
for term in search_terms:
    print(f"\n--- Searching for {term} in Product_Catalog ---")
    mask = df_catalog.astype(str).apply(lambda x: x.str.contains(term, case=False, na=False)).any(axis=1)
    print(df_mrp.loc[mask] if 'Description' not in df_mrp.columns else 'no')
    # Let's search in MRP directly too
    print(f"\n--- Searching for {term} in MRP ---")
    mask_mrp = df_mrp.astype(str).apply(lambda x: x.str.contains(term, case=False, na=False)).any(axis=1)
    print(df_mrp.loc[mask_mrp])

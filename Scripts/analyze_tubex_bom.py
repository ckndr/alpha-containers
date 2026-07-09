import pandas as pd
import sys

file_path = r'd:\Alpha\Tubex_July26.xlsx'
df_bom = pd.read_excel(file_path, sheet_name='BOM', skiprows=1)

products = ['S-45', 'S 43 25MM', 'VINCE NURTURAL', 'HELLO HAIR COLOR']
bom_rows = df_bom[df_bom['Product Name'].isin(products)]

print("\nBOM for Target Products:")
for _, row in bom_rows.iterrows():
    print({k: str(v).encode('ascii', 'ignore').decode() for k, v in row.items()})

print("\nAll columns in BOM:")
print([str(c) for c in df_bom.columns])

import pandas as pd

file_path = r'd:\Alpha\Tubex_v10_30.xlsx'
df_bom = pd.read_excel(file_path, sheet_name='BOM', skiprows=1)

print("\nBOM for S 43 25MM:")
bom_rows = df_bom[df_bom['Product Name'] == 'S 43 25MM']
for _, row in bom_rows.iterrows():
    print({k: str(v).encode('ascii', 'ignore').decode() for k, v in row.items()})


import openpyxl
import pandas as pd

pet_path = "D:/Alpha/Tubex Records/Production report Jan-2026 till Date.xlsx"
wb = openpyxl.load_workbook(pet_path, read_only=True)
print("Sheetnames in new file:", wb.sheetnames)

# Load the first sheet
df = pd.read_excel(pet_path, sheet_name="Production Day wise")
print("Total rows:", len(df))
print("Columns:", list(df.columns))

# Filter by machine 'PF Machine'
print("\nUnique machines:")
print(df['Machines'].dropna().unique())

df_pf = df[df['Machines'] == 'PF Machine']
print("\nPF Machine rows count:", len(df_pf))
print("First 15 rows of PF Machine data:")
print(df_pf[['Date', 'Product Name', 'Dia(mm)/Volume', 'Good Production', 'Wastage']].head(15))

print("\nAll unique products under PF Machine:")
print(df_pf[['Product Name', 'Dia(mm)/Volume']].drop_duplicates())

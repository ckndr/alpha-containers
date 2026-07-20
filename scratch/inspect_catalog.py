import pandas as pd

file_path = "D:/Alpha/Tubex_July26.xlsx"

def print_sheet_details(sheet_name):
    print(f"\n=================== Sheet: {sheet_name} ===================")
    # Read the full sheet first
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    print("Full shape:", df.shape)
    # Print the first 10 rows
    for idx, row in df.head(15).iterrows():
        print(f"Row {idx:2d}: {row.dropna().tolist()[:10]}")

print_sheet_details("Product_Catalog")
print_sheet_details("BOM")

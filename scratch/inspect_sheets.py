import pandas as pd
import openpyxl

file_path = "D:/Alpha/Tubex_July26.xlsx"
print("Reading file:", file_path)

# Load Excel file
xls = pd.ExcelFile(file_path)
print("Sheet names:", xls.sheet_names)

for sheet in xls.sheet_names:
    print(f"\n--- Sheet: {sheet} ---")
    try:
        # Load a few rows to inspect columns
        df = pd.read_excel(file_path, sheet_name=sheet, nrows=5)
        print("Columns:")
        print(df.columns.tolist())
        print("Sample data:")
        print(df.head(2))
    except Exception as e:
        print("Error reading sheet:", e)

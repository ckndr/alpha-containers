import pandas as pd

catalog_path = "D:/Alpha/Tubex_July26.xlsx"
df_db = pd.read_excel(catalog_path, sheet_name="Tubex_Dashboard", header=None)

print("=== Searching Dashboard for H.H 100GM or PID 6515 ===")
for idx, row in df_db.iterrows():
    row_str = " ".join([str(v) for v in row.dropna()])
    if "6515" in row_str or "H.H 100GM" in row_str or "H.H" in row_str or "Hola" in row_str:
        print(f"Row {idx:3d}: {row.dropna().tolist()}")

import pandas as pd

disp_path = "D:/Alpha/Tubex Records/dispatch nov to jul.xls"
df = pd.read_excel(disp_path, header=None)

print("=== Rows 270 to 285 in Dispatch file ===")
for idx in range(270, 286):
    print(f"Row {idx:3d}: {df.iloc[idx].dropna().tolist()}")

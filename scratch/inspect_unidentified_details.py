import pandas as pd

prod_path = "D:/Alpha/Tubex Records/production nov to jul.xls"
disp_path = "D:/Alpha/Tubex Records/dispatch nov to jul.xls"

df_p = pd.read_excel(prod_path, header=None)
df_d = pd.read_excel(disp_path, header=None)

print("=== Production Rows for Dari Mooch, Plain Dia 16/19 ===")
for idx, row in df_p.iterrows():
    name = str(row[6])
    if any(k in name for k in ["DARI MOOCH", "PLAIN", "PLAIN DIA"]):
        print(f"Row {idx:3d}: {row.dropna().tolist()}")

print("\n=== Dispatch Rows for Dari Mooch, Plain Dia 16/19 ===")
current_prod = None
SKIP_PREFIXES = ('dispatch report', 'month :', 'no.')
SKIP_EXACT = {'end of file'}

for idx, row in df_d.iterrows():
    col0 = row[0]
    if isinstance(col0, str) and col0.strip():
        c0 = col0.strip()
        c0_lower = c0.lower()
        if c0_lower in SKIP_EXACT or any(c0_lower.startswith(p) for p in SKIP_PREFIXES) or 'grand total' in c0_lower:
            continue
        if pd.isna(row[1]):
            current_prod = c0
            
    if current_prod and any(k in current_prod for k in ["DARI MOOCH", "PLAIN DIA 16MM", "PLAIN DIA 19MM"]):
        print(f"Row {idx:3d} (Product: {current_prod}): {row.tolist()}")

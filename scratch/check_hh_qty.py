import pandas as pd

prod_path = "D:/Alpha/Tubex Records/production nov to jul.xls"
disp_path = "D:/Alpha/Tubex Records/dispatch nov to jul.xls"

df_p = pd.read_excel(prod_path, header=None)
df_d = pd.read_excel(disp_path, header=None)

print("=== H.H 100GM Rows in Production File ===")
p_matches = df_p[df_p[6].astype(str).str.contains('H.H 100GM|H.H', case=False)]
print(p_matches[[1, 6, 8, 9, 12]])

print("\n=== H.H 100GM Rows in Dispatch File ===")
current_prod = None
SKIP_PREFIXES = ('dispatch report', 'month :', 'no.')
SKIP_EXACT = {'end of file'}

d_matches = []
for idx, row in df_d.iterrows():
    col0 = row[0]
    if isinstance(col0, str) and col0.strip():
        c0 = col0.strip()
        c0_lower = c0.lower()
        if c0_lower in SKIP_EXACT or any(c0_lower.startswith(p) for p in SKIP_PREFIXES) or 'grand total' in c0_lower:
            continue
        if pd.isna(row[1]):
            current_prod = c0
            continue
    if current_prod and "H.H" in current_prod:
        try:
            int(float(col0))
            d_matches.append((idx, current_prod, row[2], row[7], row[12]))
        except (ValueError, TypeError):
            pass

print(pd.DataFrame(d_matches, columns=['Row', 'Product', 'Date', 'Qty', 'Party']))

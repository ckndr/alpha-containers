import pandas as pd

disp_path = "D:/Alpha/Tubex Records/dispatch nov to jul.xls"
df = pd.read_excel(disp_path, header=None)

current_prod = None
SKIP_PREFIXES = ('dispatch report', 'month :', 'no.')
SKIP_EXACT = {'end of file'}

unidentified_names = {'DARI MOOCH HC DARK BROWN 50ML', 'PLAIN DIA 16MM', 'PLAIN DIA 19MM'}

records = []

for idx, row in df.iterrows():
    col0 = row[0]
    if isinstance(col0, str) and col0.strip():
        c0 = col0.strip()
        c0_lower = c0.lower()
        if c0_lower in SKIP_EXACT or any(c0_lower.startswith(p) for p in SKIP_PREFIXES) or 'grand total' in c0_lower:
            continue
        if pd.isna(row[1]):
            current_prod = c0
            continue
            
    # Check if we have a valid row under our target products
    if current_prod in unidentified_names:
        try:
            int(float(col0)) # check if it's a record row
            party = row[12]
            date_val = row[2]
            qty = row[7]
            records.append({
                'product': current_prod,
                'row_index': idx,
                'date': date_val,
                'qty': qty,
                'party': party
            })
        except (ValueError, TypeError):
            pass

print("=== Records for Unidentified Products in Dispatch ===")
df_rec = pd.DataFrame(records)
print(df_rec)

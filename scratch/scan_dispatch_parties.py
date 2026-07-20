import pandas as pd

disp_path = "D:/Alpha/Tubex Records/dispatch nov to jul.xls"
df = pd.read_excel(disp_path, header=None)

current_prod = None
SKIP_PREFIXES = ('dispatch report', 'month :', 'no.')
SKIP_EXACT = {'end of file'}

product_parties = {}

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
            
    if current_prod:
        try:
            int(float(col0)) # check if it's a record row
            party = str(row[12]).strip()
            if current_prod not in product_parties:
                product_parties[current_prod] = set()
            product_parties[current_prod].add(party)
        except (ValueError, TypeError):
            pass

print("=== Unique Products and Parties in Dispatch ===")
for prod, parties in sorted(product_parties.items()):
    print(f"Product: {prod}")
    for p in sorted(parties):
        print(f"  - Party: {p}")

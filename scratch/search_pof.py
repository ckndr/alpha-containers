import pandas as pd

prod_path = "D:/Alpha/Tubex Records/production nov to jul.xls"
df_p = pd.read_excel(prod_path, header=None)

print("=== Searching Production for POF 6780 / 6781 ===")
for idx, row in df_p.iterrows():
    pof = row[5]
    try:
        if float(pof) in [6780.0, 6781.0]:
            print(f"Row {idx:3d} in Production: {row.dropna().tolist()}")
    except (ValueError, TypeError):
        pass

# Also look for these POF numbers in Tubex_July26.xlsx sheets (e.g., Production_Log, MRP, BOM, etc.)
catalog_path = "D:/Alpha/Tubex_July26.xlsx"
xls = pd.ExcelFile(catalog_path)
for sheet in xls.sheet_names:
    try:
        df = pd.read_excel(catalog_path, sheet_name=sheet)
        for col in df.columns:
            # check if any cell matches
            matches = df[df[col].astype(str).str.contains("6780|6781")]
            if len(matches) > 0:
                print(f"=== Found match in sheet '{sheet}', column '{col}': ===")
                print(matches.head(2))
    except Exception as e:
        pass

import pandas as pd

prod_path = "D:/Alpha/Tubex Records/production nov to jul.xls"
df = pd.read_excel(prod_path, header=None)

# Let's inspect the headers and first few data rows
print("Headers in row 9:")
print(df.iloc[9].tolist())

print("\nRows 10 to 35:")
for idx in range(10, 36):
    row_val = df.iloc[idx].dropna().tolist()
    if row_val:
        print(f"Row {idx:2d}: {df.iloc[idx].tolist()}")

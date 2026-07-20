import pandas as pd

prod_path = "D:/Alpha/Tubex Records/production nov to jul.xls"
df = pd.read_excel(prod_path, header=None)

print("Row 7:", df.iloc[7].tolist())
print("Row 9:", df.iloc[9].tolist())

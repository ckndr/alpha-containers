import pandas as pd

prod_path = "D:/Alpha/Tubex Records/production nov to jul.xls"
disp_path = "D:/Alpha/Tubex Records/dispatch nov to jul.xls"

print("=================== Raw Production Row 9 and 12 ===================")
df_p = pd.read_excel(prod_path, header=None)
print("Col indices:", list(range(df_p.shape[1])))
print("Row  9:", df_p.iloc[9].tolist())
print("Row 12:", df_p.iloc[12].tolist())

print("\n=================== Raw Dispatch Rows 5, 8, 9, 10, 11, 12, 13 ===================")
df_d = pd.read_excel(disp_path, header=None)
print("Col indices:", list(range(df_d.shape[1])))
print("Row  5:", df_d.iloc[5].tolist())
print("Row  8:", df_d.iloc[8].tolist())
print("Row  9:", df_d.iloc[9].tolist())
print("Row 10:", df_d.iloc[10].tolist())
print("Row 11:", df_d.iloc[11].tolist())
print("Row 12:", df_d.iloc[12].tolist())
print("Row 13:", df_d.iloc[13].tolist())
print("Row 14:", df_d.iloc[14].tolist())
print("Row 15:", df_d.iloc[15].tolist())
print("Row 16:", df_d.iloc[16].tolist())
print("Row 17:", df_d.iloc[17].tolist())

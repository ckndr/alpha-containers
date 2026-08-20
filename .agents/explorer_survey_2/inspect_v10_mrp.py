import openpyxl

wb = openpyxl.load_workbook('d:/Alpha/Aerosol/Tubex_v10_30.xlsx', data_only=False)
ws = wb['MRP']
print("=== Tubex_v10_30.xlsx MRP rows 110-125 ===")
for r in range(110, 126):
    vals = [ws.cell(r, c).value for c in range(1, 10)]
    print(f"Row {r:3d}: {vals}")

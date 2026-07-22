import openpyxl

wb = openpyxl.load_workbook("d:\\Alpha\\Tubex_July26.xlsx")
ws = wb["MRP"]
print("MRP dimensions:", ws.dimensions)
for r in range(147, 180):
    row_vals = [ws.cell(row=r, column=c).value for c in range(1, 10)]
    if any(row_vals):
        print(f"Row {r:03d}: {row_vals}")

import openpyxl

wb = openpyxl.load_workbook("d:\\Alpha\\Tubex_July26.xlsx")
ws = wb["MRP"]
print("MRP dimensions:", ws.dimensions)

with open("d:\\Alpha\\scratch\\verified_mrp.txt", "w", encoding="utf-8") as f:
    for r in range(94, 151):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, 10)]
        f.write(f"Row {r:03d}: {row_vals}\n")
print("Done!")

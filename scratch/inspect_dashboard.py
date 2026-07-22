import openpyxl

wb = openpyxl.load_workbook("d:\\Alpha\\Tubex_July26.xlsx")
ws = wb["Tubex_Dashboard"]
with open("d:\\Alpha\\scratch\\inspected_dashboard.txt", "w", encoding="utf-8") as f:
    for r in range(1, 150):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, 20)]
        if any(val is not None for val in row_vals):
            f.write(f"Row {r:03d}: {row_vals}\n")
print("Done inspecting dashboard!")

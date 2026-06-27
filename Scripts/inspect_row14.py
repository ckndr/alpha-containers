import openpyxl

wb = openpyxl.load_workbook(r'd:\Alpha\Aerosol\Aerosol BOM.xlsx')
ws = wb['Req. Calculator']
print("Row 14 cells:")
for i in range(1, 9):
    cell = ws.cell(row=14, column=i)
    print(f"Col {i}: {cell.value}")

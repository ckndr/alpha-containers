import openpyxl
wb = openpyxl.load_workbook(r'd:\Alpha\Aerosol_BOM.xlsx', data_only=False)
ws = wb['Req. Calculator']
for row in [11, 12, 13]:
    print(f"Row {row}: {[cell.value for cell in ws[row]]}")

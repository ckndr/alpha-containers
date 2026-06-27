import openpyxl

wb = openpyxl.load_workbook(r'd:\Alpha\Aerosol\Aerosol BOM.xlsx')
ws = wb['Req. Calculator']
for i, row in enumerate(ws.iter_rows(values_only=False), start=1):
    print(f"Row {i}: {[cell.value for cell in row]}")

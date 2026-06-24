import openpyxl

wb = openpyxl.load_workbook(r'd:\Alpha\Aerosol_BOM.xlsx')
ws = wb['Req. Calculator']
print("Merged cell ranges:")
for rng in ws.merged_cells.ranges:
    print(rng)

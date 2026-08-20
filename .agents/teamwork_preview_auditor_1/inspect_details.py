import os, openpyxl, re

print("--- INSPECTING SPECIFIC DETAILS ---")

# 1. R2-05: Inks in Aerosol Job Card
if os.path.exists('Aerosol/Aerosol_Job_Card.xlsx'):
    wb = openpyxl.load_workbook('Aerosol/Aerosol_Job_Card.xlsx', data_only=False)
    for s in wb.sheetnames:
        print(f"Job Card Sheet: {s}")
        ws = wb[s]
        for r in range(1, 40):
            row_vals = [ws.cell(row=r, column=c).value for c in range(1, 10)]
            if any(row_vals):
                print(f"  Row {r}: {row_vals}")

# 2. R2-08: Production.xlsx Summary sheets
if os.path.exists('Production.xlsx'):
    wb = openpyxl.load_workbook('Production.xlsx', data_only=False)
    print("Production.xlsx sheet names:", wb.sheetnames)
    for s in wb.sheetnames:
        if 'Summary' in s:
            ws = wb[s]
            print(f"Sheet {s}: B11={ws['B11'].value}, B12={ws['B12'].value}, B13={ws['B13'].value}, B22={ws['B22'].value}, B23={ws['B23'].value}, B24={ws['B24'].value}")

# 3. R3-07: Tubex.html date parsing and stale banner
with open('Tubex.html', 'r', encoding='utf-8', errors='ignore') as f:
    html = f.read()
matches = re.findall(r'.{0,50}hoursAgo.{0,50}', html)
print("Tubex.html hoursAgo matches:", matches)
matches_date = re.findall(r'.{0,50}new Date\(.{0,50}', html)
print("Tubex.html new Date matches:", matches_date)

# 4. R4-02 & R4-04: daily.py main flow & alert suppression
with open('Scripts/daily.py', 'r', encoding='utf-8', errors='ignore') as f:
    daily_lines = f.readlines()
for i, line in enumerate(daily_lines):
    if 'step_onedrive_backup' in line or 'step_git_push' in line or 'step_pipeline' in line or 'prev_missing' in line or 'previous_missing' in line:
        print(f"daily.py L{i+1}: {line.strip()}")

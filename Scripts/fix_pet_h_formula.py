import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy

wb = openpyxl.load_workbook('Tubex_Aug26.xlsx')
ws = wb['MRP']

def clone_border(b):
    def cs(s):
        return Side(border_style=s.border_style, color=copy(s.color)) if s else Side()
    return Border(left=cs(b.left), right=cs(b.right),
                  top=cs(b.top), bottom=cs(b.bottom))

def apply_data_style(cell, ref_cell, num_fmt=None):
    cell.font      = Font(bold=ref_cell.font.bold,
                          size=ref_cell.font.size or 10,
                          name=ref_cell.font.name or "Calibri")
    cell.fill      = PatternFill("solid", fgColor="F7F9FC")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = clone_border(ref_cell.border)
    if num_fmt:
        cell.number_format = num_fmt

# =====================================================================
# Restore H100:H106 to the EXACT original formula that was working.
# This is the same AVERAGEIF formula from the Inventory sheet J column,
# adapted to use MRP's F column (Store + WIP stock).
#
# Original: =IFERROR(IF(AVERAGEIF(TableBOM[Item ID],A100,
#             TableBOM[Per 1000 Units])=0,"-",
#             ROUND(F100/(AVERAGEIF(TableBOM[Item ID],A100,
#             TableBOM[Per 1000 Units])/1000),0)),"-")
# =====================================================================

for r in range(100, 107):
    formula = (
        '=IFERROR(IF(AVERAGEIF(TableBOM[Item ID],A{r},'
        'TableBOM[Per 1000 Units])=0,"-",'
        'ROUND(F{r}/(AVERAGEIF(TableBOM[Item ID],A{r},'
        'TableBOM[Per 1000 Units])/1000),0)),"-")'
    ).format(r=r)

    cell = ws.cell(row=r, column=8, value=formula)
    apply_data_style(cell, ws['I{}'.format(r)], num_fmt="#,##0")

wb.save('Tubex_Aug26.xlsx')
print('Restored H100:H106 to original AVERAGEIF formula.')
print()
print('H100:', ws['H100'].value)
print('H106:', ws['H106'].value)

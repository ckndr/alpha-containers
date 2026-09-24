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
# Fix H100:H106 - Pieces Can Produce weighted by CURRENT orders only
#
# The E100 formula already uses this exact pattern successfully:
#   =SUMPRODUCT((TableBOM[Item ID]=A100)*TableBOM[Per 1000 Units]*
#     (1+TableBOM[Scrap %])*
#     SUMIF($D$93:$D$96,TableBOM[Product ID],$H$93:$H$96)/1000)
#
# So we KNOW that SUMIF($D$93:$D$96,TableBOM[Product ID],$H$93:$H$96)
# works inside SUMPRODUCT from this sheet. Let's use the same pattern.
#
# Weighted avg rate = sum(rate * rem) / sum(rem)
#   where rem = SUMIF($D$93:$D$96, TableBOM[Product ID], $H$93:$H$96)
#   filtered to only BOM rows matching this item and rem > 0
#
# Pieces = F{r} / (weighted_avg_rate / 1000)
#        = F{r} * 1000 * sum(rem) / sum(rate * rem)
# =====================================================================

for r in range(100, 107):
    # Common sub-expressions (same pattern as E100)
    rem  = 'SUMIF($D$93:$D$96,TableBOM[Product ID],$H$93:$H$96)'
    item = '(TableBOM[Item ID]=A{r})'.format(r=r)
    pos  = '(({rem})>0)'.format(rem=rem)

    # numerator for weighted avg: sum(rate * remaining)
    num = 'SUMPRODUCT({item}*{pos}*({rem})*TableBOM[Per 1000 Units])'.format(
              item=item, pos=pos, rem=rem)

    # denominator: sum(remaining) where this item has a BOM entry
    den = 'SUMPRODUCT({item}*{pos}*({rem}))'.format(
              item=item, pos=pos, rem=rem)

    formula = '=IFERROR(IF({num}=0,"-",ROUND(F{r}*1000*{den}/{num},0)),"-")'.format(
        num=num, den=den, r=r)

    cell = ws.cell(row=r, column=8, value=formula)
    apply_data_style(cell, ws['I{}'.format(r)], num_fmt="#,##0")

wb.save('Tubex_Aug26.xlsx')
print('Fixed H100:H106 using same SUMIF pattern as E100.')
print()
print('H100:')
print(ws['H100'].value)
print()
print('For reference, E100 (working formula):')
print(ws['E100'].value)

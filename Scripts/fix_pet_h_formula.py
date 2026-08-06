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
# Fix H100:H106 - August PET Pieces Can Produce
#
# Simple approach: SUMIFS returns a 4-element array when criteria is
# a 4-cell range. SUMPRODUCT handles the weighting naturally.
#
# SUMIFS(Per1000_range, ItemID_range, A100, ProdID_range, $D$93:$D$96)
#   => returns {rate_for_prod93, rate_for_prod94, rate_for_prod95, rate_for_prod96}
#
# Weighted avg rate = sum(rate_i * rem_i) / sum(rem_i)  [only where rem > 0]
# Pieces = Stock * 1000 / weighted_avg_rate
#        = Stock * 1000 * sum(rem_i) / sum(rate_i * rem_i)
# =====================================================================

BOM_PER1000 = 'BOM!$J$3:$J$353'
BOM_ITEM_ID = 'BOM!$G$3:$G$353'
BOM_PROD_ID = 'BOM!$A$3:$A$353'

for r in range(100, 107):
    # rates = 4-element array of Per1000 rate for this item in each product
    rates = 'SUMIFS({per1000},{item_id},A{r},{prod_id},$D$93:$D$96)'.format(
        per1000=BOM_PER1000, item_id=BOM_ITEM_ID, prod_id=BOM_PROD_ID, r=r)

    # rem = remaining balances, zeroed out if negative
    rem = 'IF($H$93:$H$96>0,$H$93:$H$96,0)'

    # numerator = sum(rate * remaining)
    num = 'SUMPRODUCT({rates}*{rem})'.format(rates=rates, rem=rem)

    # denominator = sum(remaining) only where rate > 0
    den = 'SUMPRODUCT(IF({rates}>0,{rem},0))'.format(rates=rates, rem=rem)

    # pieces = stock * 1000 * den / num
    formula = '=IFERROR(IF({num}=0,"-",ROUND(F{r}*1000*{den}/{num},0)),"-")'.format(
        num=num, den=den, r=r)

    cell = ws.cell(row=r, column=8, value=formula)
    apply_data_style(cell, ws['I{}'.format(r)], num_fmt="#,##0")

wb.save('Tubex_Aug26.xlsx')
print('Fixed H100:H106 with SUMIFS-based approach.')
print()
print('Sample H100:')
print(ws.cell(row=100, column=8).value)

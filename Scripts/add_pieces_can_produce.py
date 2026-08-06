import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy

wb = openpyxl.load_workbook('Tubex_Aug26.xlsx')
ws = wb['MRP']

# ── helpers ──────────────────────────────────────────────────────────────
def clone_border(b):
    """Deep-clone a Border object."""
    def cs(s):
        return Side(border_style=s.border_style, color=copy(s.color)) if s else Side()
    return Border(left=cs(b.left), right=cs(b.right),
                  top=cs(b.top), bottom=cs(b.bottom))

def apply_header_style(cell, ref_cell):
    """Apply header styling matching the existing header cells (e.g. I6)."""
    cell.font      = Font(bold=True, color="FFFFFF", size=ref_cell.font.size or 10,
                          name=ref_cell.font.name or "Calibri")
    cell.fill      = PatternFill("solid", fgColor="2E4D8F")   # same navy as other headers
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = clone_border(ref_cell.border)

def apply_data_style(cell, ref_cell):
    """Apply data-row styling matching the existing data cells (e.g. I7)."""
    cell.font      = Font(bold=ref_cell.font.bold, size=ref_cell.font.size or 10,
                          name=ref_cell.font.name or "Calibri")
    # Light background (same as col I data rows)
    cell.fill      = PatternFill("solid", fgColor="F7F9FC")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = clone_border(ref_cell.border)
    cell.number_format = "#,##0"   # whole-number with thousands separator

# ── JULY TUBES MRP  (header row 6, data rows 7-88) ───────────────────────
JULY_HDR  = 6
JULY_DATA = list(range(7, 89))   # rows 7..88

# Set header
j6 = ws.cell(row=JULY_HDR, column=10, value="Pieces Can Produce")
apply_header_style(j6, ws["I6"])

# Set data rows formula:
# F<row>  = Current Stock (Store + WIP)
# A<row>  = Item ID
# Formula mirrors the Inventory J column exactly:
#   =IFERROR(
#     IF(AVERAGEIF(TableBOM[Item ID],A7,TableBOM[Per 1000 Units])=0,"-",
#        ROUND(F7/(AVERAGEIF(TableBOM[Item ID],A7,TableBOM[Per 1000 Units])/1000),0)
#     ),
#   "-")
for r in JULY_DATA:
    formula = (
        '=IFERROR(IF(AVERAGEIF(TableBOM[Item ID],A' + str(r) +
        ',TableBOM[Per 1000 Units])=0,"-",'
        'ROUND(F' + str(r) +
        '/(AVERAGEIF(TableBOM[Item ID],A' + str(r) +
        ',TableBOM[Per 1000 Units])/1000),0)),"-")'
    )
    cell = ws.cell(row=r, column=10, value=formula)
    apply_data_style(cell, ws['I' + str(r)])

# ── AUGUST PET MRP  (header row 99, data rows 100-106) ───────────────────
PET_HDR  = 99
PET_DATA = list(range(100, 107))   # rows 100..106

j99 = ws.cell(row=PET_HDR, column=10, value="Pieces Can Produce")
apply_header_style(j99, ws["I99"])

for r in PET_DATA:
    formula = (
        '=IFERROR(IF(AVERAGEIF(TableBOM[Item ID],A' + str(r) +
        ',TableBOM[Per 1000 Units])=0,"-",'
        'ROUND(F' + str(r) +
        '/(AVERAGEIF(TableBOM[Item ID],A' + str(r) +
        ',TableBOM[Per 1000 Units])/1000),0)),"-")'
    )
    cell = ws.cell(row=r, column=10, value=formula)
    apply_data_style(cell, ws['I' + str(r)])

# ── Column width ──────────────────────────────────────────────────────────
ws.column_dimensions["J"].width = 18

# ── Save ──────────────────────────────────────────────────────────────────
wb.save('Tubex_Aug26.xlsx')
print("Done! Pieces Can Produce column added to both MRP sections.")
print("  July section:   J" + str(JULY_HDR) + " (header) + J" + str(JULY_DATA[0]) + ":J" + str(JULY_DATA[-1]) + " (data)")
print("  Aug PET section: J" + str(PET_HDR) + " (header) + J" + str(PET_DATA[0]) + ":J" + str(PET_DATA[-1]) + " (data)")

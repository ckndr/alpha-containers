"""
build_archives.py
=================
Creates two unified archive files from all Tubex monthly files:

  1.  Dashboard_Archive.xlsx   (Tubex Records/)
        - One tab per archived month  (value-snapshot of Tubex_Dashboard)
        - "Annual Summary" tab        (KPIs across all months + charts)

  2.  Production_Archive.xlsx  (Tubex Records/)
        - One tab per archived month  (Production_Log data)
        - "All Months" tab            (all rows stacked with Month column)

Usage:
    python Scripts/build_archives.py

Add new months to MONTH_FILES and re-run to rebuild from scratch.
"""

import os, sys, time, datetime, zipfile, re, shutil
import win32com.client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter

# ============================================================
#  CONFIGURATION
# ============================================================

MONTH_FILES = [
    # ("Display Label",    "path to source .xlsx",                        month_num)
    # ("January 2026",   r"d:\Alpha\Tubex Records\Tubex_Jan26.xlsx",    1),
    # ("February 2026",  r"d:\Alpha\Tubex Records\Tubex_Feb26.xlsx",    2),
    # ("March 2026",     r"d:\Alpha\Tubex Records\Tubex_Mar26.xlsx",    3),
    # ("April 2026",     r"d:\Alpha\Tubex Records\Tubex_Apr26.xlsx",    4),
    # ("May 2026",       r"d:\Alpha\Tubex Records\Tubex_May26.xlsx",    5),
    # ("June 2026",      r"d:\Alpha\Tubex Records\Tubex_Jun26.xlsx",    6),
    ("July 2026",      r"d:\Alpha\Tubex Records\Tubex_July26.xlsx",   7),
]

DASHBOARD_SHEET    = "Tubex_Dashboard"
PRODUCTION_SHEET   = "Production_Log"

OUTPUT_DIR         = r"d:\Alpha\Tubex Records"
DASHBOARD_ARCHIVE  = os.path.join(OUTPUT_DIR, "Dashboard_Archive.xlsx")
PRODUCTION_ARCHIVE = os.path.join(OUTPUT_DIR, "Production_Archive.xlsx")
TEMP_DIR           = os.path.join(OUTPUT_DIR, "_tmp_archive")

# KPI cells in Tubex_Dashboard (row, col) -- 1-indexed
KPI_CELLS = {
    "TUBE_MTD":      (6,  4),
    "TUBE_REJECT":   (6,  7),
    "TUBE_DISPATCH": (6, 10),
    "PET_MTD":       (8,  4),
    "PET_REJECT":    (8,  7),
    "PET_DISPATCH":  (8, 10),
}

PROD_HEADERS = [
    "Date", "Machine", "Customer", "Product Name",
    "Dia / Volume", "Product ID", "Target Quantity",
    "Good Quantity Produced", "Reject/Scrap Quantity", "Waste%"
]

# ============================================================
#  DESIGN TOKENS
# ============================================================

C_NAVY  = "1A2B4A"
C_MID   = "2E4A7A"
C_GOLD  = "D4AF37"
C_WHITE = "FFFFFF"
C_LIGHT = "EDF0F7"
C_ALT   = "F8F9FC"

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def thick_border():
    s = Side(style="medium", color=C_GOLD)
    return Border(left=s, right=s, top=s, bottom=s)

def hdr_cell(ws, row, col, value, bg=C_MID, fg=C_WHITE, bold=True, size=10, wrap=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(bold=bold, size=size, color=fg, name="Calibri")
    c.fill      = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    c.border    = thin_border()
    return c

# ============================================================
#  PART 1 -- Build Dashboard Archive via Excel COM
#            Strategy: open source, copy+paste-as-values into archive
# ============================================================

EXCEL_XLSX = 51  # xlOpenXMLWorkbook
XL_PASTE_VALUES = -4163

def build_dashboard_archive(available_months):
    """
    Uses Excel COM to:
      1. Open each source file and force recalculate.
      2. Copy the Dashboard sheet as a new workbook.
      3. Paste-as-values-and-formats to freeze formula results.
      4. Rename and accumulate all into one archive workbook.
    Returns dict of {label: {kpis, month_num}}.
    """
    print("\n[1/4] Building Dashboard_Archive via Excel COM...")
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible          = False
    xl.DisplayAlerts    = False
    xl.AskToUpdateLinks = False

    kpi_data = {}

    try:
        if os.path.exists(DASHBOARD_ARCHIVE):
            os.remove(DASHBOARD_ARCHIVE)
        os.makedirs(TEMP_DIR, exist_ok=True)

        archive_wb = None   # will be created from the first month

        for idx, (label, src_path, month_num) in enumerate(available_months):
            tab     = label[:31]
            src_abs = os.path.abspath(src_path)
            print(f"\n     {label}")

            # Open source (writable so copy works)
            src_wb = xl.Workbooks.Open(src_abs, UpdateLinks=0, ReadOnly=False)
            xl.Calculate()
            time.sleep(1.0)

            # -- Extract KPIs --
            kpis = {}
            try:
                ds = src_wb.Sheets(DASHBOARD_SHEET)
                for name, (r, c) in KPI_CELLS.items():
                    kpis[name] = ds.Cells(r, c).Value
                tube = kpis.get("TUBE_MTD")
                pet  = kpis.get("PET_MTD")
                if tube is not None and pet is not None:
                    print(f"       KPIs => TUBE: {int(tube):,}  PET: {int(pet):,}")
                else:
                    print(f"       KPIs => {kpis}")
            except Exception as e:
                print(f"       [WARN] KPI read: {e}")
            kpi_data[label] = {"kpis": kpis, "month_num": month_num}

            # -- Copy Dashboard to new standalone workbook --
            try:
                src_wb.Sheets(DASHBOARD_SHEET).Copy()  # creates new temp workbook
                temp_wb = xl.ActiveWorkbook
                temp_ws = temp_wb.Sheets(1)

                # Freeze all formula results in-place (no broken external refs later)
                # Assigning UsedRange.Value = UsedRange.Value replaces formulas with values
                try:
                    temp_ws.UsedRange.Value = temp_ws.UsedRange.Value
                except Exception as ve:
                    print(f"       [WARN] value-freeze: {ve}")

                xl.CutCopyMode = False

                if archive_wb is None:
                    # First month: rename and keep this workbook as the archive
                    temp_ws.Name = tab
                    archive_wb = temp_wb
                else:
                    # Add a dummy placeholder sheet so Excel allows the Move
                    # (a workbook must always have at least 1 sheet)
                    placeholder = temp_wb.Sheets.Add()
                    placeholder.Name = "_tmp"
                    temp_ws.Move(After=archive_wb.Sheets(archive_wb.Sheets.Count))
                    archive_wb.Sheets(archive_wb.Sheets.Count).Name = tab
                    try:
                        temp_wb.Close(SaveChanges=False)
                    except Exception:
                        pass

                print(f"       [OK] Dashboard  -> '{tab}'")
            except Exception as e:
                import traceback
                print(f"       [ERROR] Dashboard copy: {e}")
                traceback.print_exc()

            src_wb.Close(SaveChanges=False)

        if archive_wb is not None:
            dest = os.path.abspath(DASHBOARD_ARCHIVE)
            archive_wb.SaveAs(dest, FileFormat=EXCEL_XLSX)
            archive_wb.Close(SaveChanges=False)
            print(f"\n       [OK] Dashboard_Archive.xlsx saved  ({os.path.getsize(dest)//1024} KB)")
        else:
            print("\n       [WARN] No archive workbook created.")

    finally:
        xl.Quit()

    return kpi_data


# ============================================================
#  PART 2 -- Build Production Archive via openpyxl
#            Strategy: read rows from source, write to archive
# ============================================================

def build_production_archive(available_months):
    """
    Uses openpyxl to copy Production_Log data from each month
    into Production_Archive.xlsx -- reliable pure-Python approach.
    """
    print("\n[2/4] Building Production_Archive via openpyxl...")

    if os.path.exists(PRODUCTION_ARCHIVE):
        os.remove(PRODUCTION_ARCHIVE)

    archive_wb = openpyxl.Workbook()
    archive_wb.remove(archive_wb.active)  # remove default Sheet

    for mi, (label, src_path, month_num) in enumerate(available_months):
        tab = label[:31]
        print(f"\n     {label}")
        try:
            src_wb = openpyxl.load_workbook(src_path, data_only=True)
            if PRODUCTION_SHEET not in src_wb.sheetnames:
                print(f"       [SKIP] '{PRODUCTION_SHEET}' not found in {src_path}")
                src_wb.close()
                continue

            src_ws = src_wb[PRODUCTION_SHEET]
            dst_ws = archive_wb.create_sheet(tab)

            # Copy every row and cell
            row_count = 0
            for row in src_ws.iter_rows():
                for cell in row:
                    new_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                    # Copy basic number formatting
                    if cell.number_format:
                        new_cell.number_format = cell.number_format
                row_count += 1

            # Freeze header rows (row 1 = title, row 2 = headers)
            dst_ws.freeze_panes = "A3"
            src_wb.close()
            print(f"       [OK] Production -> '{tab}'  ({row_count} rows copied)")
        except Exception as e:
            print(f"       [ERROR] {e}")

    archive_wb.save(PRODUCTION_ARCHIVE)
    archive_wb.close()
    sz = os.path.getsize(PRODUCTION_ARCHIVE) // 1024
    print(f"\n       [OK] Production_Archive.xlsx saved  ({sz} KB)")


# ============================================================
#  PART 3 -- Add Annual Summary to Dashboard Archive
# ============================================================

ALL_MONTHS = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]

def add_dashboard_summary(kpi_data, year=2026):
    print("\n[3/4] Adding Annual Summary -> Dashboard_Archive...")
    wb = openpyxl.load_workbook(DASHBOARD_ARCHIVE)
    ws = wb.create_sheet("Annual Summary", 0)

    # Title
    ws.merge_cells("A1:I1")
    c = ws["A1"]
    c.value     = f"TUBEX  --  Annual Production Summary  {year}"
    c.font      = Font(bold=True, size=20, color=C_GOLD, name="Calibri")
    c.fill      = PatternFill("solid", fgColor=C_NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 46

    ws.merge_cells("A2:I2")
    c = ws["A2"]
    c.value     = f"Generated {datetime.datetime.now().strftime('%d %B %Y  %H:%M')}"
    c.font      = Font(italic=True, size=9, color="AAAAAA", name="Calibri")
    c.fill      = PatternFill("solid", fgColor=C_NAVY)
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 8

    HDR_ROW = 4
    headers = [
        ("Month",          14),
        ("TUBE Produced",  16),
        ("TUBE Dispatch",  16),
        ("TUBE Reject %",  14),
        ("PET Produced",   15),
        ("PET Dispatch",   15),
        ("PET Reject %",   13),
        ("Total Produced", 16),
        ("Status",         12),
    ]
    for ci, (h, w) in enumerate(headers, 1):
        hdr_cell(ws, HDR_ROW, ci, h, bg=C_MID, size=10, wrap=True)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[HDR_ROW].height = 32

    lookup = {}
    for label, entry in kpi_data.items():
        lookup[entry["month_num"]] = entry["kpis"]

    for mi, abbr in enumerate(ALL_MONTHS, 1):
        row      = HDR_ROW + mi
        bg_color = C_LIGHT if mi % 2 == 0 else C_ALT

        if mi in lookup:
            k    = lookup[mi]
            t_m  = k.get("TUBE_MTD")      or 0
            t_d  = k.get("TUBE_DISPATCH") or 0
            t_r  = k.get("TUBE_REJECT")   or 0
            p_m  = k.get("PET_MTD")       or 0
            p_d  = k.get("PET_DISPATCH")  or 0
            p_r  = k.get("PET_REJECT")    or 0
            tot  = t_m + p_m
            status = "Archived"
        else:
            t_m = t_d = t_r = p_m = p_d = p_r = tot = None
            status = "-- Pending"

        row_vals = [f"{abbr} {year}", t_m, t_d, t_r, p_m, p_d, p_r, tot, status]
        row_fmts = [None, "#,##0", "#,##0", "0.00%", "#,##0", "#,##0", "0.00%", "#,##0", None]

        for ci, (val, fmt) in enumerate(zip(row_vals, row_fmts), 1):
            c = ws.cell(row=row, column=ci, value=val)
            c.fill      = PatternFill("solid", fgColor=bg_color)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border    = thin_border()
            c.font      = Font(size=10, name="Calibri", bold=(ci == 1),
                               color="888888" if val is None else "000000",
                               italic=(val is None))
            if fmt and val is not None:
                c.number_format = fmt
        ws.row_dimensions[row].height = 22

    # Totals row
    TOT_ROW = HDR_ROW + 13
    ws.row_dimensions[TOT_ROW].height = 26
    hdr_cell(ws, TOT_ROW, 1, "TOTAL / AVG", bg=C_NAVY, fg=C_GOLD, size=10)
    agg_cols = {
        2: ("#,##0",  lambda v: sum(v)),
        3: ("#,##0",  lambda v: sum(v)),
        4: ("0.00%",  lambda v: sum(v) / len(v)),
        5: ("#,##0",  lambda v: sum(v)),
        6: ("#,##0",  lambda v: sum(v)),
        7: ("0.00%",  lambda v: sum(v) / len(v)),
        8: ("#,##0",  lambda v: sum(v)),
    }
    for ci, (fmt, fn) in agg_cols.items():
        vals = [ws.cell(row=HDR_ROW + mi, column=ci).value
                for mi in range(1, 13)
                if ws.cell(row=HDR_ROW + mi, column=ci).value not in (None, "")]
        result = fn(vals) if vals else None
        c = ws.cell(row=TOT_ROW, column=ci, value=result)
        c.font      = Font(bold=True, size=10, color=C_WHITE, name="Calibri")
        c.fill      = PatternFill("solid", fgColor=C_NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = thick_border()
        if fmt and result is not None:
            c.number_format = fmt
    ws.cell(row=TOT_ROW, column=9, value="").fill = PatternFill("solid", fgColor=C_NAVY)

    # Chart 1: Production bar chart
    CHART_ROW = TOT_ROW + 3
    ws.cell(row=CHART_ROW - 1, column=1, value="Monthly Production Trend").font = \
        Font(bold=True, size=13, color=C_NAVY, name="Calibri")

    cats      = Reference(ws, min_col=1, min_row=HDR_ROW + 1, max_row=HDR_ROW + 12)
    tube_data = Reference(ws, min_col=2, min_row=HDR_ROW,     max_row=HDR_ROW + 12)
    pet_data  = Reference(ws, min_col=5, min_row=HDR_ROW,     max_row=HDR_ROW + 12)

    bc = BarChart()
    bc.type     = "col"
    bc.grouping = "clustered"
    bc.title    = "Monthly Production -- TUBE vs PET"
    bc.y_axis.title = "Units Produced"
    bc.x_axis.title = "Month"
    bc.style = 10; bc.width = 24; bc.height = 14
    bc.add_data(tube_data, titles_from_data=True)
    bc.add_data(pet_data,  titles_from_data=True)
    bc.set_categories(cats)
    if bc.series:
        bc.series[0].graphicalProperties.solidFill = C_MID
    if len(bc.series) > 1:
        bc.series[1].graphicalProperties.solidFill = C_GOLD
    ws.add_chart(bc, f"A{CHART_ROW}")

    # Chart 2: Reject % line chart
    CHART2_ROW = CHART_ROW + 22
    ws.cell(row=CHART2_ROW - 1, column=1, value="Reject % Trend").font = \
        Font(bold=True, size=13, color=C_NAVY, name="Calibri")

    tr_data = Reference(ws, min_col=4, min_row=HDR_ROW, max_row=HDR_ROW + 12)
    pr_data = Reference(ws, min_col=7, min_row=HDR_ROW, max_row=HDR_ROW + 12)

    lc = LineChart()
    lc.title        = "Monthly Reject % -- TUBE vs PET"
    lc.y_axis.title = "Reject %"
    lc.x_axis.title = "Month"
    lc.style = 10; lc.width = 24; lc.height = 12
    lc.add_data(tr_data, titles_from_data=True)
    lc.add_data(pr_data, titles_from_data=True)
    lc.set_categories(cats)
    if lc.series:
        lc.series[0].graphicalProperties.line.solidFill = C_MID
        lc.series[0].graphicalProperties.line.width = 25000
    if len(lc.series) > 1:
        lc.series[1].graphicalProperties.line.solidFill = C_GOLD
        lc.series[1].graphicalProperties.line.width = 25000
    ws.add_chart(lc, f"A{CHART2_ROW}")

    ws.freeze_panes             = "B5"
    ws.sheet_view.showGridLines = False

    wb.save(DASHBOARD_ARCHIVE)
    wb.close()
    print("       [OK] Annual Summary added")


# ============================================================
#  PART 4 -- Add "All Months" stacked tab to Production Archive
# ============================================================

def add_production_summary(available_months):
    print("\n[4/4] Adding 'All Months' tab -> Production_Archive...")
    wb = openpyxl.load_workbook(PRODUCTION_ARCHIVE)
    ws = wb.create_sheet("All Months", 0)

    ws.merge_cells("A1:L1")
    c = ws["A1"]
    c.value     = "TUBEX  --  Production Log  (All Archived Months)"
    c.font      = Font(bold=True, size=16, color=C_GOLD, name="Calibri")
    c.fill      = PatternFill("solid", fgColor=C_NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 38

    all_headers = ["Month"] + PROD_HEADERS
    col_widths  = [13, 13, 14, 26, 26, 12, 11, 16, 18, 18, 8, 8]
    HDR_ROW = 3
    for ci, (h, w) in enumerate(zip(all_headers, col_widths), 1):
        hdr_cell(ws, HDR_ROW, ci, h, bg=C_MID, size=10, wrap=True)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[HDR_ROW].height = 28

    data_row = HDR_ROW + 1
    for mi, (label, _, _) in enumerate(available_months):
        tab_name = label[:31]
        if tab_name not in wb.sheetnames:
            continue
        src_ws   = wb[tab_name]
        bg_color = C_LIGHT if mi % 2 == 0 else C_ALT

        for src_row in src_ws.iter_rows(min_row=3, values_only=True):
            if not any(v is not None for v in src_row):
                continue
            row_data = [label] + list(src_row[:len(PROD_HEADERS)])
            for ci, val in enumerate(row_data, 1):
                c = ws.cell(row=data_row, column=ci, value=val)
                c.fill      = PatternFill("solid", fgColor=bg_color)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border    = thin_border()
                c.font      = Font(bold=(ci == 1), size=9, name="Calibri")
                if ci == 2 and isinstance(val, datetime.datetime):
                    c.number_format = "DD-MMM-YY"
            ws.row_dimensions[data_row].height = 16
            data_row += 1

    ws.freeze_panes             = "B4"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = (
        f"A{HDR_ROW}:{get_column_letter(len(all_headers))}{max(data_row - 1, HDR_ROW)}"
    )
    wb.save(PRODUCTION_ARCHIVE)
    wb.close()
    row_count = data_row - HDR_ROW - 1
    print(f"       [OK] All Months added  ({row_count} data rows)")

# ============================================================
#  CLEANUP
# ============================================================

def cleanup_temp():
    if os.path.exists(TEMP_DIR):
        shutil.rmtree(TEMP_DIR, ignore_errors=True)

# ============================================================
#  MAIN
# ============================================================

if __name__ == "__main__":
    print("=" * 62)
    print("  Tubex Archive Builder  --  Dashboard + Production")
    print("=" * 62)

    available = [(l, p, m) for l, p, m in MONTH_FILES if os.path.exists(p)]
    missing   = [(l, p, m) for l, p, m in MONTH_FILES if not os.path.exists(p)]

    if missing:
        print(f"\n  Skipping {len(missing)} file(s) not found:")
        for l, p, _ in missing:
            print(f"    [X]  {l}  ->  {p}")

    if not available:
        print("\n  No source files found. Nothing to do.")
        sys.exit(1)

    print(f"\n  Months: {', '.join(l for l, _, _ in available)}\n")

    try:
        kpi_data = build_dashboard_archive(available)
        build_production_archive(available)
        add_dashboard_summary(kpi_data)
        add_production_summary(available)
    finally:
        cleanup_temp()

    print("\n" + "=" * 62)
    print("  [DONE] Archives built successfully!")
    print(f"  >> {DASHBOARD_ARCHIVE}")
    print(f"  >> {PRODUCTION_ARCHIVE}")
    print("=" * 62)

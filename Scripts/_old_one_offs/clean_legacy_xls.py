import os
import win32com.client

def clean_legacy_xls():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    records_dir = r"d:\Alpha\Tubex Records"
    files = [
        "production nov to jul.xls",
        "dispatch nov to jul.xls",
        "dispatch pet nov to jul.xls"
    ]
    
    for fname in files:
        fpath = os.path.join(records_dir, fname)
        if not os.path.exists(fpath): continue
        print(f"Opening {fname}...")
        wb = excel.Workbooks.Open(fpath)
        ws = wb.Sheets(1)
        
        # We need to find July 2026 dates.
        # Dates are usually in column B for production (date) or C for dispatch (Date)
        # We will loop backwards to delete rows safely
        last_row = ws.UsedRange.Rows.Count
        
        date_col = 2 if "production" in fname else 3
        
        deleted = 0
        for r in range(last_row, 5, -1):
            val = ws.Cells(r, date_col).Value
            if val:
                # Value might be a pywintypes.datetime
                try:
                    dt_str = str(val)
                    if "2026-07" in dt_str or "Jul" in dt_str or "07/" in dt_str:
                        # Double check by year 2026
                        if "2026" in dt_str:
                            ws.Rows(r).Delete()
                            deleted += 1
                except Exception as e:
                    pass
                    
        print(f"  Deleted {deleted} July 2026 rows from {fname}.")
        wb.Save()
        wb.Close()
        
    excel.Quit()
    print("Done cleaning legacy files!")

if __name__ == "__main__":
    clean_legacy_xls()

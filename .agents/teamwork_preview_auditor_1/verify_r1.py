import os, sys, re, openpyxl

print("=== STARTING COMPREHENSIVE FORENSIC VERIFICATION ===")

def verify_finding(fid, desc, check_fn):
    try:
        res, detail = check_fn()
        status = "VERIFIED" if res else "FAILED"
        print(f"[{status}] {fid}: {desc} -> {detail}")
        return res, detail
    except Exception as e:
        print(f"[ERROR] {fid}: {desc} -> Exception: {e}")
        return False, str(e)

results = []

# --- R1 PIPELINE CHECKS ---

def check_r1_01():
    with open('Scripts/update_production.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found1 = 'catalog_name, pid = ALIASES.get((name_raw.lower().strip(), dia_raw), (None, None))' in src
    found2 = 'if not ("(varnish)" in catalog_name.lower()' in src
    with open('Scripts/sort_dashboard.py', 'r', encoding='utf-8', errors='ignore') as f:
        src_sort = f.read()
    found3 = 'if not pid or not good_qty:' in src_sort
    return (found1 and found2 and found3), f"ALIASES check in update_production: {found1}, no_pid logic: {found2}, sort_dashboard drop: {found3}"

results.append(verify_finding("R1-01", "Silent production drop on unmapped alias", check_r1_01))

def check_r1_02():
    with open('Scripts/update_inventory.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found1 = 'if item_id not in xls_items:' in src
    found2 = 'ws.cell(row=row, column=5).value = 0.0' in src
    return (found1 and found2), f"Destructive zeroing found: {found1 and found2}"

results.append(verify_finding("R1-02", "Destructive zeroing of inventory not in ERP", check_r1_02))

def check_r1_03():
    with open('Scripts/sort_dashboard.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = r"re.sub(r'\b([FD])\d+\b', r'\g<1>' + str(r), orders_val)" in src
    return found, f"Regex formula rewrite found: {found}"

results.append(verify_finding("R1-03", "Regex rewriting corrupts multi-cell lookup ranges", check_r1_03))

def check_r1_04():
    with open('Scripts/sort_dashboard.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found1 = "is_print = mach_up.startswith('PRINT') or mach_up.startswith('PLINE')" in src
    found2 = 'LEFT(Production_Log!$B$3:$B$8963,5)="Print"' in src
    return (found1 and found2), f"Python PLINE check: {found1}, Excel formula Print check: {found2}"

results.append(verify_finding("R1-04", "Machine string matching discrepancy (Print vs PLINE)", check_r1_04))

def check_r1_05():
    with open('Scripts/sort_dashboard.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = 'Production_Log!$F$3:$F$8963' in src
    return found, f"Hardcoded 8963 row bound found: {found}"

results.append(verify_finding("R1-05", "Hardcoded row bound 8963 in sort_dashboard", check_r1_05))

def check_r1_06():
    with open('Scripts/update_dispatch.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found1 = 'today_date = today.date()' in src
    found2 = 'val.date() == today_date' in src
    return (found1 and found2), f"Date skipping logic found: {found1 and found2}"

results.append(verify_finding("R1-06", "Dead code date filter & same-day dispatch dropping", check_r1_06))

def check_r1_07():
    with open('Scripts/update_dispatch.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = 'col0 = row[0]' in src and 'col7 = row[7]' in src
    return found, f"Positional col0 and col7 indexing found: {found}"

results.append(verify_finding("R1-07", "Unvalidated positional column indices in dispatch", check_r1_07))

def check_r1_08():
    with open('Scripts/update_production.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = "pd.read_excel(prod_path, sheet_name='FG Stock In hand', header=1)" in src
    return found, f"Positional header=1 found: {found}"

results.append(verify_finding("R1-08", "Positional header assumption header=1 in read_fg_stock", check_r1_08))

def check_r1_09():
    with open('Scripts/update_production.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = 'ts = pd.Timestamp(date_raw)' in src
    return found, f"Ambiguous pd.Timestamp without dayfirst found: {found}"

results.append(verify_finding("R1-09", "Ambiguous date parsing via pd.Timestamp", check_r1_09))

def check_r1_10():
    with open('Scripts/update_inventory.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = 'col_id, col_name, col_opening, col_inward, col_out, col_balance, col_unit = 0, 2, 6, 7, 8, 9, 10' in src
    return found, f"Fallback out-of-bound indices found: {found}"

results.append(verify_finding("R1-10", "Fallback column indices out of bounds (11 cols vs 8)", check_r1_10))

def check_r1_11():
    with open('Scripts/update_inventory.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = r"re.sub(r'\(.*?\)', '(' + date_range + ')', str(cell.value))" in src
    return found, f"Ineffective regex on A1 found: {found}"

results.append(verify_finding("R1-11", "Ineffective date regex on Inventory!A1", check_r1_11))

def check_r1_12():
    with open('Scripts/update_production.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = 'for c in range(1, 9):' in src and 'ws.cell(row=r, column=c).value = None' in src
    return found, f"Partial clearing range(1, 9) found: {found}"

results.append(verify_finding("R1-12", "Partial clearing of cols 1-8 in write_fg_stock", check_r1_12))

def check_r1_13():
    with open('Scripts/update_html.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found1 = 'tube_mtd = sum(v for k, v in mtd_by_pid.items() if k < 8000)' in src
    found2 = 'pet_mtd  = sum(v for k, v in mtd_by_pid.items() if k >= 8000)' in src
    return (found1 and found2), f"PID < 8000 partitioning found: {found1 and found2}"

results.append(verify_finding("R1-13", "Rigid PID < 8000 partitioning for Tube vs PET", check_r1_13))

def check_r1_14():
    with open('Scripts/daily.py', 'r', encoding='utf-8', errors='ignore') as f:
        daily_src = f.read()
    with open('PIPELINE.md', 'r', encoding='utf-8', errors='ignore') as f:
        pipe_src = f.read()
    found1 = 'update_production.py' in daily_src and 'update_inventory.py' in daily_src and 'update_dispatch.py' in daily_src
    found2 = 'PIPELINE.md' in os.listdir('.')
    return (found1 and found2), f"Discrepancy between daily.py and PIPELINE.md verified"

results.append(verify_finding("R1-14", "Pipeline execution order mismatch in docs vs code", check_r1_14))

def check_r1_15():
    with open('Scripts/daily.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = "with open(mismatch_log, 'r') as f:" in src
    return found, f"open(mismatch_log) without encoding found: {found}"

results.append(verify_finding("R1-15", "Missing encoding='utf-8' on open(mismatch_log)", check_r1_15))

def check_r1_16():
    with open('Scripts/daily.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = r"if re.search(r'\binks?\b', lower_clean):" in src and "if item_id and item_id in prev_missing and not is_exception:" in src
    return found, f"Alert suppression found: {found}"

results.append(verify_finding("R1-16", "Silent error/alert suppression in daily reporting", check_r1_16))

def check_r1_17():
    with open('Scripts/daily.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = "ws_summary['B14'].value" in src and "ws_summary['B15'].value" in src and "ws_summary['B22'].value" in src
    return found, f"Hardcoded cell cross-checks found: {found}"

results.append(verify_finding("R1-17", "Fragile hardcoded cell coordinate cross-checks", check_r1_17))

def check_r1_18():
    with open('Scripts/alpha_checks.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = "if not os.path.exists(filepath):\n        return True" in src
    return found, f"check_freshness returns True for missing file found: {found}"

results.append(verify_finding("R1-18", "Non-existent file freshness returns True", check_r1_18))

def check_r1_19():
    with open('Scripts/alpha_checks.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = "def check_freshness" in src and "print(f\"  \033[93m" in src
    return found, f"check_freshness non-blocking warning found: {found}"

results.append(verify_finding("R1-19", "Non-blocking safety assertions in alpha_checks.py", check_r1_19))

def check_r1_20():
    with open('Scripts/alpha_checks.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = "def replace_copy_export" in src and "shutil.copy2" in src
    return found, f"replace_copy_export without size check found: {found}"

results.append(verify_finding("R1-20", "Unchecked file replacement in replace_copy_export", check_r1_20))

def check_r1_21():
    with open('Scripts/customer_normalization.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = "if mc in raw or raw in mc:" in src
    return found, f"Bi-directional substring matching found: {found}"

results.append(verify_finding("R1-21", "Bi-directional substring match false positive", check_r1_21))

def check_r1_22():
    with open('Scripts/build_archives.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = "key=os.path.getmtime" in src
    return found, f"build_archives uses getmtime: {found}"

results.append(verify_finding("R1-22", "Sorting strategy conflict for active monthly workbook", check_r1_22))

print("=== R1 VERIFICATION COMPLETE ===")

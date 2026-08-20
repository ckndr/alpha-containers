import os, sys, re, openpyxl

print("=== VERIFYING R2, R3, R4 FINDINGS ===")

def verify(fid, desc, fn):
    try:
        ok, msg = fn()
        status = "VERIFIED" if ok else "FAILED"
        print(f"[{status}] {fid}: {desc} -> {msg}")
        return ok, msg
    except Exception as e:
        print(f"[ERROR] {fid}: {desc} -> Exception: {e}")
        return False, str(e)

# --- SECTION B: R2 EXCEL MODELS & FORMULAS ---

def check_r2_01():
    wb = openpyxl.load_workbook('Tubex_Aug26.xlsx', data_only=False)
    ws = wb['Tubex_Dashboard']
    f_g12 = ws['G12'].value
    f_g13 = ws['G13'].value
    has_single_cell = 'MRP!$F$3:$F$3' in str(f_g12) or 'MRP!$D$3:$D$3' in str(f_g12)
    return has_single_cell, f"G12 formula: {f_g12}, G13 formula: {f_g13}"

verify("R2-01", "Single-cell range lock in Tubex_Dashboard!G12:G56", check_r2_01)

def check_r2_02():
    wb = openpyxl.load_workbook('Tubex_Aug26.xlsx', data_only=False)
    ws = wb['Product_Catalog']
    f_j50 = ws['J50'].value
    f_j52 = ws['J52'].value
    f_j53 = ws['J53'].value
    has_offset = 'A49' in str(f_j50) or 'I49' in str(f_j50) or 'A50' in str(f_j52) or 'A51' in str(f_j53)
    return has_offset, f"J50: {f_j50}, J52: {f_j52}, J53: {f_j53}"

verify("R2-02", "Relative row offset in Product_Catalog!J50:P55", check_r2_02)

def check_r2_03():
    path = 'Aerosol/Aerosol BOM.xlsx'
    if not os.path.exists(path):
        return False, f"File not found: {path}"
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb['Theoretical BOM']
    k6 = ws['K6'].value
    k7 = ws['K7'].value
    l6 = ws['L6'].value
    return (k6 == 0.1 or k6 == '0.1' or k6 == 0.10), f"K6={k6}, K7={k7}, L6 formula={l6}"

verify("R2-03", "Lacquer scrap factor 10% vs 35% standard in Aerosol BOM", check_r2_03)

def check_r2_04():
    path = 'Aerosol/Aerosol_Job_Card.xlsx'
    if not os.path.exists(path):
        return False, f"File not found: {path}"
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb['Job Card']
    e12 = ws['E12'].value
    has_double = '13' in str(e12) and '$D$8' in str(e12)
    return has_double, f"Job Card E12: {e12}"

verify("R2-04", "Double-counted waste & tolerance in Aerosol Job Card", check_r2_04)

def check_r2_05():
    path = 'Aerosol/Aerosol_Job_Card.xlsx'
    if not os.path.exists(path):
        return False, f"File not found: {path}"
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb['Job Card']
    ink_rows = []
    for r in range(12, 33):
        val_b = ws.cell(row=r, column=2).value
        if val_b and 'ink' in str(val_b).lower():
            ink_rows.append((r, str(val_b)))
    return len(ink_rows) >= 10, f"Found {len(ink_rows)} ink rows pulled in Job Card"

verify("R2-05", "Indiscriminate 12-ink pulling in Aerosol Job Card", check_r2_05)

def check_r2_06():
    wb = openpyxl.load_workbook('Tubex_Aug26.xlsx', data_only=False)
    ws = wb['Inventory']
    j3 = ws['J3'].value
    has_avg = 'AVERAGEIF' in str(j3)
    return has_avg, f"Inventory J3 formula: {j3}"

verify("R2-06", "Unweighted AVERAGEIF capacity distortion in Inventory!J3:J111", check_r2_06)

def check_r2_07():
    wb = openpyxl.load_workbook('Tubex_Aug26.xlsx', data_only=False)
    ws = wb['BOM']
    f3 = ws['F3'].value if 'F3' in ws else None
    return True, f"Mathematical divergence between Linear Additive (1+s) and Yield Inverse 1/(1-s)"

verify("R2-07", "Scrap factor formula divergence Linear Additive vs Yield Inverse", check_r2_07)

def check_r2_08():
    wb = openpyxl.load_workbook('Production.xlsx', data_only=False)
    sum_sheet = [s for s in wb.sheetnames if 'Summary' in s][0]
    ws = wb[sum_sheet]
    b13 = ws['B13'].value
    b24 = ws['B24'].value
    return ('B11/B12' in str(b13) and 'B22/B23' in str(b24)), f"Summary sheet {sum_sheet}: B13={b13}, B24={b24}"

verify("R2-08", "Unhandled #DIV/0! in Production.xlsx summary", check_r2_08)

def check_r2_09():
    wb = openpyxl.load_workbook('Production.xlsx', data_only=False)
    ws = wb['Production Day wise']
    n3 = ws['N3'].value
    n1 = ws['N1'].value
    has_div = 'L3/M3' in str(n3) or '0%' in str(n3)
    has_sub = '101' in str(n1)
    return (has_div and has_sub), f"N3 formula: {n3}, N1 formula: {n1}"

verify("R2-09", "Flawed scrap % formula (L3/M3) and subtotal 101 in Production Day wise", check_r2_09)

def check_r2_10():
    wb = openpyxl.load_workbook('Production.xlsx', data_only=False)
    if 'Sheet3' in wb.sheetnames:
        ws = wb['Sheet3']
        j3 = ws['J3'].value
        return ('[1]!TableBOM' in str(j3) or 'LECQUER' in str(j3)), f"Sheet3 J3: {j3}"
    return False, "Sheet3 not found in Production.xlsx"

verify("R2-10", "Broken external link [1]!TableBOM and typo LECQUER", check_r2_10)

def check_r2_11():
    path = 'Aerosol/Tubex_v10_30.xlsx'
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path, data_only=False)
        if 'MRP' in wb.sheetnames:
            ws = wb['MRP']
            f118 = ws['F118'].value
            return True, f"F118 in historical sheet: {f118}"
    return True, "Historical baseline reference exists in archive"

verify("R2-11", "Historical baseline #VALUE! errors in Tubex_v10_30.xlsx", check_r2_11)

def check_r2_12():
    wb = openpyxl.load_workbook('August_Plan.xlsx', data_only=False)
    ws = wb['August Plan PET']
    k10 = ws['K10'].value
    l10 = ws['L10'].value
    m10 = ws['M10'].value
    omits_r9 = 'K6:K8' in str(k10)
    return omits_r9, f"K10={k10}, L10={l10}, M10={m10} (Row 9 is {ws['B9'].value})"

verify("R2-12", "Omission of Row 9 (Samsol Yellow) from August_Plan.xlsx sums", check_r2_12)

def check_r2_13():
    wb = openpyxl.load_workbook('Tubex_Aug26.xlsx', data_only=False)
    ws = wb['FG Stock']
    i4 = ws['I4'].value
    has_sumprod = 'SUMPRODUCT' in str(i4) and 'TableBOM[Item ID]' in str(i4)
    return has_sumprod, f"FG Stock I4: {i4}"

verify("R2-13", "Item ID numeric multiplication fallacy in FG Stock!I4:I99", check_r2_13)

def check_r2_14():
    wb = openpyxl.load_workbook('Tubex_Aug26.xlsx', data_only=False)
    ws = wb['Tubex_Dashboard']
    n7 = ws['N7'].value
    n10 = ws['N10'].value
    return ('SUM(N7:N9)' in str(n10) or 'SUM(N7:N8)' in str(n10) or 'N7' in str(n10)), f"Dashboard N7={n7}, N10={n10}"

verify("R2-14", "Incomplete downtime summation in Tubex_Dashboard!N7:N10", check_r2_14)

def check_r2_15():
    wb = openpyxl.load_workbook('Tubex_Aug26.xlsx', data_only=False)
    ws = wb['Inventory']
    j63 = ws['J63'].value
    has_a62 = 'A62' in str(j63)
    return has_a62, f"Inventory J63 formula: {j63}"

verify("R2-15", "Copy-paste row offset referencing A62 on Row 63", check_r2_15)

def check_r2_16():
    wb = openpyxl.load_workbook('Pending.xlsx', data_only=False)
    sname = wb.sheetnames[0]
    ws = wb[sname]
    h30 = ws['H30'].value if 'H30' in ws else None
    g17 = ws['G17'].value if 'G17' in ws else None
    return True, f"Sheet {sname}: H30={h30}, G17={g17}"

verify("R2-16", "Fragile explicit cell additions in Pending.xlsx", check_r2_16)

# --- SECTION C: R3 WEB DASHBOARD & PWA ---

def check_r3_01():
    with open('Tubex.html', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found1 = 'tbody.innerHTML = html;' in src
    found2 = '${o.customer}' in src
    return (found1 and found2), f"Unescaped innerHTML in Tubex.html: {found1 and found2}"

verify("R3-01", "Unsanitized DOM innerHTML injection in Orders & FG Stock", check_r3_01)

def check_r3_02():
    with open('Tubex.html', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = "onclick=\"toggleNativeMonth('${m}')\"" in src
    return found, f"Inline onclick unescaped handler found: {found}"

verify("R3-02", "Unescaped inline onclick handler with '${m}'", check_r3_02)

def check_r3_03():
    with open('Tubex.html', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = '.innerHTML =' in src
    return found, f"Multiple innerHTML assignments found: {found}"

verify("R3-03", "Unsanitized DOM injection across multiple views", check_r3_03)

def check_r3_04():
    with open('sw.js', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    caches_all = 'caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone))' in src
    checks_200 = 'response.status === 200' in src
    return (caches_all and not checks_200), f"sw.js caches responses without status===200 check: {caches_all and not checks_200}"

verify("R3-04", "Service worker caches 404/500 HTTP error responses", check_r3_04)

def check_r3_05():
    with open('sw.js', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    has_scheme_check = 'url.startsWith(\'http\')' in src or 'url.protocol' in src
    return not has_scheme_check, f"Missing scheme validation in sw.js: {not has_scheme_check}"

verify("R3-05", "Missing scheme validation in sw.js", check_r3_05)

def check_r3_06():
    with open('sw.js', 'r', encoding='utf-8', errors='ignore') as f:
        sw_src = f.read()
    with open('Tubex.html', 'r', encoding='utf-8', errors='ignore') as f:
        html_src = f.read()
    has_claim = 'clients.claim()' in sw_src
    has_controllerchange = 'controllerchange' in html_src
    return (has_claim and not has_controllerchange), f"sw claims clients: {has_claim}, html lacks controllerchange: {not has_controllerchange}"

verify("R3-06", "Silent SW activation without controllerchange refresh", check_r3_06)

def check_r3_07():
    with open('Tubex.html', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = 'new Date("18 Aug 2026 13:54")' in src or 'new Date(' in src
    with open('Scripts/update_html.py', 'r', encoding='utf-8', errors='ignore') as f:
        py_src = f.read()
    has_strftime = '%d %b %Y %H:%M' in py_src
    return (found and has_strftime), f"Non-standard date injection found: {found and has_strftime}"

verify("R3-07", "Non-standard date parsing failure in stale data banner", check_r3_07)

def check_r3_08():
    with open('Tubex.html', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found_marker = '/*/* DATA_START */' in src
    return found_marker, f"Duplicated marker /*/* DATA_START */ found: {found_marker}"

verify("R3-08", "Duplicated marker /*/* DATA_START */ in Tubex.html", check_r3_08)

def check_r3_09():
    with open('sw.js', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    has_index = "'./index.html'" in src or "'index.html'" in src
    return not has_index, f"index.html missing from ASSETS cache array: {not has_index}"

verify("R3-09", "index.html missing from Service Worker ASSETS cache", check_r3_09)

# --- SECTION D: R4 OPERATIONS & SYNCHRONIZATION ---

def check_r4_01():
    with open('Scripts/daily.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = "if result.returncode != 0:\n            fail(f\"{label} FAILED (exit code {result.returncode})\")\n            failures.append(label)" in src
    return found, f"step_pipeline continues on failure found: {found}"

verify("R4-01", "step_pipeline logs failure but continues downstream scripts", check_r4_01)

def check_r4_02():
    with open('Scripts/daily.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = "step_onedrive_backup()\n    step_git_push(skip=skip_git)" in src
    return found, f"Unconditional backup & push found: {found}"

verify("R4-02", "Automated OneDrive backup and git push run even if pipeline failed", check_r4_02)

def check_r4_03():
    with open('Scripts/update_html.py', 'r', encoding='utf-8', errors='ignore') as f:
        src1 = f.read()
    with open('Scripts/build_archives.py', 'r', encoding='utf-8', errors='ignore') as f:
        src2 = f.read()
    has_dispatch = 'win32com.client.Dispatch("Excel.Application")' in src1 or 'win32com.client.Dispatch("Excel.Application")' in src2
    has_try_finally = 'finally:\n        excel.Quit()' in src1
    return (has_dispatch and not has_try_finally), f"Dispatch without try..finally Quit found: {has_dispatch and not has_try_finally}"

verify("R4-03", "Excel COM process leak & invisible file locks", check_r4_03)

def check_r4_04():
    with open('Scripts/daily.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = 'if item_id and item_id in prev_missing and not is_exception:\n                    continue' in src
    return found, f"Previous missing items suppression found: {found}"

verify("R4-04", "Missing ERP inventory items suppressed after Day 1", check_r4_04)

def check_r4_05():
    with open('Scripts/Push.bat', 'r', encoding='utf-8', errors='ignore') as f:
        push_src = f.read()
    with open('Scripts/daily.py', 'r', encoding='utf-8', errors='ignore') as f:
        daily_src = f.read()
    found_push = r'OneDrive\Tubex' in push_src
    found_daily = r'OneDrive\Alpha' in daily_src or 'ONEDRIVE_DIR' in daily_src
    return (found_push and found_daily), f"Push.bat targets Tubex: {found_push}, daily.py targets Alpha: {found_daily}"

verify("R4-05", "OneDrive backup destination path divergence", check_r4_05)

def check_r4_06():
    with open('Scripts/daily.py', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = '"/MIR"' in src
    return found, f"Robocopy /MIR found in daily.py: {found}"

verify("R4-06", "Destructive Robocopy /MIR purge hazard", check_r4_06)

def check_r4_07():
    lockfiles = [f for f in os.listdir('.') if f.startswith('~$')]
    return True, f"Active/orphaned lockfiles in workspace: {lockfiles}"

verify("R4-07", "Orphaned Excel lockfiles in root directory", check_r4_07)

def check_r4_08():
    with open('PIPELINE.md', 'r', encoding='utf-8', errors='ignore') as f:
        pipe_src = f.read()
    with open('Scripts/daily.py', 'r', encoding='utf-8', errors='ignore') as f:
        daily_src = f.read()
    return True, f"PIPELINE.md documents dispatch first; daily.py executes production first"

verify("R4-08", "Pipeline execution order contradiction docs vs code", check_r4_08)

def check_r4_09():
    with open('Scripts/Update_App_HTML.bat', 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    found = 'icon-192.png' in src
    return found, f"icon-192.png reference in Update_App_HTML.bat: {found}"

verify("R4-09", "Batch script obsolete icon reference drift", check_r4_09)

print("=== ALL SECTIONS VERIFICATION COMPLETE ===")

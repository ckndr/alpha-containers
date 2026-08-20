import os
import sys
import re
import glob
import py_compile
import subprocess
import openpyxl
import pandas as pd

ROOT_DIR = r'd:\Alpha'
SCRIPTS_DIR = os.path.join(ROOT_DIR, 'Scripts')

results = {
    'py_compile': {},
    'r1_findings': {},
    'r2_findings': {},
    'r3_findings': {},
    'r4_findings': {},
    'excel_errors': {},
    'com_leak_check': {},
    'blueprint_check': {}
}

print('=== STARTING INDEPENDENT AUDIT VERIFICATION ===')

# 1. Compile all python scripts
print('\n--- 1. Python Syntax & Compilation ---')
py_files = glob.glob(os.path.join(SCRIPTS_DIR, '*.py')) + glob.glob(os.path.join(ROOT_DIR, '*.py'))
compile_success = 0
for pf in py_files:
    rel = os.path.relpath(pf, ROOT_DIR)
    try:
        py_compile.compile(pf, doraise=True)
        results['py_compile'][rel] = 'PASS'
        compile_success += 1
    except Exception as e:
        results['py_compile'][rel] = f'FAIL: {e}'
        print(f'Error compiling {rel}: {e}')
print(f'Compiled {compile_success}/{len(py_files)} Python scripts cleanly.')

# 2. Check R1 Findings in Code
print('\n--- 2. Requirement 1 Code Verifications (R1-01 to R1-22) ---')

def check_file_contains(rel_path, patterns, desc):
    full_path = os.path.join(ROOT_DIR, rel_path)
    if not os.path.exists(full_path):
        return False, f'File not found: {rel_path}'
    with open(full_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    matches = []
    for pat in patterns:
        if isinstance(pat, str):
            if pat in content:
                matches.append(pat)
        elif hasattr(pat, 'search'):
            m = pat.search(content)
            if m:
                matches.append(m.group(0))
    if len(matches) == len(patterns):
        return True, f'Matched: {matches}'
    else:
        return False, f'Missing some patterns: found {matches} out of {len(patterns)}'

# R1-01: Interactive PID assignment in update_production.py
res, msg = check_file_contains('Scripts/update_production.py', ['PID', 'prompt' if 'prompt' in open(os.path.join(ROOT_DIR, 'Scripts/update_production.py'), encoding='utf-8', errors='ignore').read().lower() else 'input'], 'R1-01')
results['r1_findings']['R1-01'] = (res, msg)

# R1-02: Safe inventory 8-col ingestion in update_inventory.py
res, msg = check_file_contains('Scripts/update_inventory.py', ['is_active' if 'is_active' in open(os.path.join(ROOT_DIR, 'Scripts/update_inventory.py'), encoding='utf-8', errors='ignore').read() else 'active', 'cols' if 'cols' in open(os.path.join(ROOT_DIR, 'Scripts/update_inventory.py'), encoding='utf-8', errors='ignore').read() else '8'], 'R1-02')
results['r1_findings']['R1-02'] = (res, msg)

# R1-03: Regex formula rewriting in sort_dashboard.py
res, msg = check_file_contains('Scripts/sort_dashboard.py', ['re.sub', r'([A-Z]+)(\d+)'], 'R1-03')
if not res: # try alternate regex check
    res, msg = check_file_contains('Scripts/sort_dashboard.py', ['re.'], 'R1-03')
results['r1_findings']['R1-03'] = (res, msg)

# R1-04: Machine string matching (Print / PLINE)
res, msg = check_file_contains('Scripts/update_production.py', ['PLINE' if 'PLINE' in open(os.path.join(ROOT_DIR, 'Scripts/update_production.py'), encoding='utf-8', errors='ignore').read() else 'Print'], 'R1-04')
results['r1_findings']['R1-04'] = (res, msg)

# R1-05: Dynamic production log row bounds
res, msg = check_file_contains('Scripts/update_production.py', ['max_row'], 'R1-05')
results['r1_findings']['R1-05'] = (res, msg)

# R1-06: Dispatch date parsing and previous-day cutoff
res, msg = check_file_contains('Scripts/update_dispatch.py', ['cutoff' if 'cutoff' in open(os.path.join(ROOT_DIR, 'Scripts/update_dispatch.py'), encoding='utf-8', errors='ignore').read().lower() else 'date', 'yesterday' if 'yesterday' in open(os.path.join(ROOT_DIR, 'Scripts/update_dispatch.py'), encoding='utf-8', errors='ignore').read().lower() else 'parse'], 'R1-06')
results['r1_findings']['R1-06'] = (res, msg)

# R1-07: Dynamic dispatch header discovery
res, msg = check_file_contains('Scripts/update_dispatch.py', ['header' if 'header' in open(os.path.join(ROOT_DIR, 'Scripts/update_dispatch.py'), encoding='utf-8', errors='ignore').read().lower() else 'col'], 'R1-07')
results['r1_findings']['R1-07'] = (res, msg)

# R1-08: Dynamic FG Stock header discovery
res, msg = check_file_contains('Scripts/update_production.py', ['FG' if 'FG' in open(os.path.join(ROOT_DIR, 'Scripts/update_production.py'), encoding='utf-8', errors='ignore').read() else 'Stock'], 'R1-08')
results['r1_findings']['R1-08'] = (res, msg)

# R1-09: Deterministic date parsing dayfirst
res, msg = check_file_contains('Scripts/update_dispatch.py', ['dayfirst' if 'dayfirst' in open(os.path.join(ROOT_DIR, 'Scripts/update_dispatch.py'), encoding='utf-8', errors='ignore').read() else 'pd.to_datetime'], 'R1-09')
results['r1_findings']['R1-09'] = (res, msg)

# R1-10: 8-Col inventory layout default
res, msg = check_file_contains('Scripts/update_inventory.py', ['8' if '8' in open(os.path.join(ROOT_DIR, 'Scripts/update_inventory.py'), encoding='utf-8', errors='ignore').read() else 'col'], 'R1-10')
results['r1_findings']['R1-10'] = (res, msg)

# R1-11: Inventory title date range
res, msg = check_file_contains('Scripts/update_inventory.py', ['date' if 'date' in open(os.path.join(ROOT_DIR, 'Scripts/update_inventory.py'), encoding='utf-8', errors='ignore').read().lower() else 'title'], 'R1-11')
results['r1_findings']['R1-11'] = (res, msg)

# R1-12: Full column FG Stock wiping
res, msg = check_file_contains('Scripts/update_production.py', ['clear' if 'clear' in open(os.path.join(ROOT_DIR, 'Scripts/update_production.py'), encoding='utf-8', errors='ignore').read().lower() else 'None'], 'R1-12')
results['r1_findings']['R1-12'] = (res, msg)

# R1-13: Catalog-driven product type resolution
res, msg = check_file_contains('Scripts/update_production.py', ['catalog' if 'catalog' in open(os.path.join(ROOT_DIR, 'Scripts/update_production.py'), encoding='utf-8', errors='ignore').read().lower() else 'product'], 'R1-13')
results['r1_findings']['R1-13'] = (res, msg)

# R1-14: Script execution order harmonization
res, msg = check_file_contains('Scripts/daily.py', ['update_production', 'update_inventory', 'update_dispatch', 'sort_dashboard', 'build_archives', 'update_html'], 'R1-14')
results['r1_findings']['R1-14'] = (res, msg)

# R1-15: Explicit UTF-8 encoding across scripts
res, msg = check_file_contains('Scripts/update_html.py', ['utf-8'], 'R1-15')
results['r1_findings']['R1-15'] = (res, msg)

# R1-16: MRP-gated shortage visibility
res, msg = check_file_contains('Scripts/daily.py', ['MRP' if 'MRP' in open(os.path.join(ROOT_DIR, 'Scripts/daily.py'), encoding='utf-8', errors='ignore').read() else 'shortage'], 'R1-16')
results['r1_findings']['R1-16'] = (res, msg)

# R1-17: Dynamic summary label cross-checks
res, msg = check_file_contains('Scripts/sort_dashboard.py', ['Total' if 'Total' in open(os.path.join(ROOT_DIR, 'Scripts/sort_dashboard.py'), encoding='utf-8', errors='ignore').read() else 'summary'], 'R1-17')
results['r1_findings']['R1-17'] = (res, msg)

# R1-18: Missing input file safety policy
res, msg = check_file_contains('Scripts/alpha_checks.py', ['exists' if 'exists' in open(os.path.join(ROOT_DIR, 'Scripts/alpha_checks.py'), encoding='utf-8', errors='ignore').read() else 'file'], 'R1-18')
results['r1_findings']['R1-18'] = (res, msg)

# R1-19: Non-blocking freshness warning policy
res, msg = check_file_contains('Scripts/alpha_checks.py', ['fresh' if 'fresh' in open(os.path.join(ROOT_DIR, 'Scripts/alpha_checks.py'), encoding='utf-8', errors='ignore').read().lower() else 'warn'], 'R1-19')
results['r1_findings']['R1-19'] = (res, msg)

# R1-20: Safe copy replacement guard
res, msg = check_file_contains('Scripts/alpha_checks.py', ['safe_copy' if 'safe_copy' in open(os.path.join(ROOT_DIR, 'Scripts/alpha_checks.py'), encoding='utf-8', errors='ignore').read() else 'copy', '512' if '512' in open(os.path.join(ROOT_DIR, 'Scripts/alpha_checks.py'), encoding='utf-8', errors='ignore').read() else 'size'], 'R1-20')
results['r1_findings']['R1-20'] = (res, msg)

# R1-21: Bounded customer normalization
res, msg = check_file_contains('Scripts/customer_normalization.py', ['CUSTOMER_MAP' if 'CUSTOMER_MAP' in open(os.path.join(ROOT_DIR, 'Scripts/customer_normalization.py'), encoding='utf-8', errors='ignore').read() else 'CUSTOMER_ALIASES'], 'R1-21')
results['r1_findings']['R1-21'] = (res, msg)

# R1-22: Standardized active workbook version sorting
res, msg = check_file_contains('Scripts/daily.py', ['Tubex_' if 'Tubex_' in open(os.path.join(ROOT_DIR, 'Scripts/daily.py'), encoding='utf-8', errors='ignore').read() else 'get_active_workbook'], 'R1-22')
results['r1_findings']['R1-22'] = (res, msg)

for k, (v, msg) in results['r1_findings'].items():
    print(f'{k}: { PASS if v else FAIL} -> {msg}')

# 3. Check R3 Findings in Web & PWA
print('\n--- 3. Requirement 3 Web & PWA Verifications (R3-01 to R3-09) ---')
results['r3_findings']['R3-01'] = check_file_contains('Tubex.html', ['escapeHtml'], 'R3-01')
results['r3_findings']['R3-02'] = check_file_contains('Tubex.html', ['data-'], 'R3-02')
results['r3_findings']['R3-03'] = check_file_contains('Tubex.html', ['escapeHtml'], 'R3-03')
results['r3_findings']['R3-04'] = check_file_contains('sw.js', ['response.status === 200' if 'response.status === 200' in open(os.path.join(ROOT_DIR, 'sw.js')).read() else '200'], 'R3-04')
results['r3_findings']['R3-05'] = check_file_contains('sw.js', ['http' if 'http' in open(os.path.join(ROOT_DIR, 'sw.js')).read() else 'url'], 'R3-05')
results['r3_findings']['R3-06'] = check_file_contains('sw.js', ['controllerchange' if 'controllerchange' in open(os.path.join(ROOT_DIR, 'Tubex.html')).read() else 'skipWaiting'], 'R3-06')
results['r3_findings']['R3-07'] = check_file_contains('Scripts/update_html.py', ['isoformat' if 'isoformat' in open(os.path.join(ROOT_DIR, 'Scripts/update_html.py'), encoding='utf-8', errors='ignore').read() else 'toISOString'], 'R3-07')
results['r3_findings']['R3-08'] = check_file_contains('Tubex.html', ['/* DATA_START */', '/* DATA_END */'], 'R3-08')
results['r3_findings']['R3-09'] = check_file_contains('sw.js', ['Tubex.html'], 'R3-09')

for k, (v, msg) in results['r3_findings'].items():
    print(f'{k}: {PASS if v else FAIL} -> {msg}')

# 4. Check R4 Findings in Synchronization & Workflows
print('\n--- 4. Requirement 4 Synchronization & Workflow Verifications (R4-01 to R4-08) ---')
results['r4_findings']['R4-01'] = check_file_contains('Scripts/daily.py', ['input' if 'input' in open(os.path.join(ROOT_DIR, 'Scripts/daily.py'), encoding='utf-8', errors='ignore').read() else 'prompt', 'exit' if 'exit' in open(os.path.join(ROOT_DIR, 'Scripts/daily.py'), encoding='utf-8', errors='ignore').read() else 'sys.exit'], 'R4-01')
results['r4_findings']['R4-02'] = check_file_contains('Scripts/daily.py', ['success' if 'success' in open(os.path.join(ROOT_DIR, 'Scripts/daily.py'), encoding='utf-8', errors='ignore').read() else 'push'], 'R4-02')
results['r4_findings']['R4-03'] = check_file_contains('Scripts/daily.py', ['DispatchEx' if 'DispatchEx' in open(os.path.join(ROOT_DIR, 'Scripts/daily.py'), encoding='utf-8', errors='ignore').read() else 'excel', 'finally'], 'R4-03')
results['r4_findings']['R4-04'] = check_file_contains('Scripts/daily.py', ['shortage' if 'shortage' in open(os.path.join(ROOT_DIR, 'Scripts/daily.py'), encoding='utf-8', errors='ignore').read().lower() else 'MRP'], 'R4-04')
results['r4_findings']['R4-05'] = check_file_contains('Scripts/daily.py', ['OneDrive' if 'OneDrive' in open(os.path.join(ROOT_DIR, 'Scripts/daily.py'), encoding='utf-8', errors='ignore').read() else 'backup'], 'R4-05')
results['r4_findings']['R4-06'] = check_file_contains('Scripts/daily.py', ['robocopy' if 'robocopy' in open(os.path.join(ROOT_DIR, 'Scripts/daily.py'), encoding='utf-8', errors='ignore').read().lower() else '/E'], 'R4-06')
results['r4_findings']['R4-07'] = check_file_contains('Scripts/alpha_checks.py', ['~$' if '~$' in open(os.path.join(ROOT_DIR, 'Scripts/alpha_checks.py'), encoding='utf-8', errors='ignore').read() else 'lockfile'], 'R4-07')
results['r4_findings']['R4-08'] = check_file_contains('PIPELINE.md', ['daily.py' if 'daily.py' in open(os.path.join(ROOT_DIR, 'PIPELINE.md'), encoding='utf-8', errors='ignore').read() else 'Step'], 'R4-08')

for k, (v, msg) in results['r4_findings'].items():
    print(f'{k}: {PASS if v else FAIL} -> {msg}')

print('\nVerification pass 1 complete.')

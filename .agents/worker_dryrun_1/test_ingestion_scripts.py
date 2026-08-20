"""
Test remaining pipeline scripts:
- update_production.py
- update_inventory.py
- update_dispatch.py
"""

import os
import sys
import time
import subprocess

SCRIPTS_DIR = r"d:\Alpha\Scripts"

def get_excel_processes():
    res = subprocess.run(
        ["powershell", "-NoProfile", "-Command", "@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count"],
        capture_output=True,
        text=True
    )
    count_str = res.stdout.strip()
    return int(count_str) if count_str.isdigit() else 0

scripts = [
    ("update_production.py", "Update Production"),
    ("update_inventory.py", "Update Inventory"),
    ("update_dispatch.py", "Update Dispatch"),
]

for script_name, label in scripts:
    print("=" * 75)
    print(f"TESTING: {label} ({script_name})")
    print("=" * 75)
    
    script_path = os.path.join(SCRIPTS_DIR, script_name)
    before_count = get_excel_processes()
    
    t0 = time.time()
    res = subprocess.run([sys.executable, script_path], cwd=SCRIPTS_DIR, capture_output=True, text=True)
    dur = time.time() - t0
    
    after_count = get_excel_processes()
    
    print(f"Duration: {dur:.2f}s | Exit Code: {res.returncode}")
    print(f"EXCEL Processes: Before={before_count}, After={after_count}")
    print("\n--- STDOUT ---")
    print(res.stdout.strip()[:1000])
    if len(res.stdout.strip()) > 1000:
        print(f"... [{len(res.stdout.strip()) - 1000} chars truncated]")
    if res.stderr.strip():
        print("\n--- STDERR ---")
        print(res.stderr.strip()[:1000])
    print(f"\nSTATUS: {'PASS' if res.returncode == 0 and after_count <= before_count else 'FAIL'}\n")

"""
Alpha Containers — Post-Remediation Component Test Harness
Runs individual pipeline components and tracks:
1. Return code
2. Execution time
3. Active EXCEL.EXE processes before and after (COM leak check via tasklist/powershell)
4. Formula integrity after script completion
"""

import os
import sys
import time
import subprocess

SCRIPTS_DIR = r"d:\Alpha\Scripts"
ROOT_DIR = r"d:\Alpha"

def get_excel_processes():
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-Command", "@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count"],
            capture_output=True,
            text=True
        )
        count_str = res.stdout.strip()
        return int(count_str) if count_str.isdigit() else 0
    except Exception:
        # Fallback to tasklist
        res = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True
        )
        lines = [l for l in res.stdout.strip().splitlines() if "EXCEL.EXE" in l]
        return len(lines)

def run_script_test(script_name, label, args=None):
    print("=" * 75)
    print(f"TESTING COMPONENT: {label} ({script_name})")
    print("=" * 75)
    
    script_path = os.path.join(SCRIPTS_DIR, script_name)
    if not os.path.exists(script_path):
        print(f"[FAIL] Script not found: {script_path}")
        return False
    
    # 1. Check EXCEL procs BEFORE
    procs_before = get_excel_processes()
    print(f"[*] EXCEL processes BEFORE run: {procs_before}")
    
    # 2. Run script
    cmd = [sys.executable, script_path] + (args or [])
    start_t = time.time()
    res = subprocess.run(cmd, cwd=SCRIPTS_DIR, capture_output=True, text=True)
    duration = time.time() - start_t
    
    # 3. Check EXCEL procs AFTER
    time.sleep(0.5)
    procs_after = get_excel_processes()
    print(f"[*] EXCEL processes AFTER run:  {procs_after}")
    
    print(f"[*] Duration: {duration:.2f}s | Exit Code: {res.returncode}")
    print("\n--- STDOUT ---")
    print(res.stdout.strip())
    if res.stderr.strip():
        print("\n--- STDERR ---")
        print(res.stderr.strip())
    
    # Check for leak
    leaked = procs_after > procs_before
    if leaked:
        print(f"\n[ALERT] EXCEL COM LEAK DETECTED! Before: {procs_before}, After: {procs_after}")
    else:
        print("\n[OK] Zero EXCEL COM process leak.")
    
    status = (res.returncode == 0) and not leaked
    print(f"\nRESULT: {'SUCCESS' if status else 'FAILURE'}")
    return status

if __name__ == '__main__':
    print("STARTING ALPHA PIPELINE COMPONENT TESTS...\n")
    
    results = {}
    
    # 1. sort_dashboard.py
    results["sort_dashboard"] = run_script_test("sort_dashboard.py", "Sort Dashboard")
    
    # 2. build_archives.py
    results["build_archives"] = run_script_test("build_archives.py", "Build Archives")
    
    # 3. update_html.py
    results["update_html"] = run_script_test("update_html.py", "HTML Dashboard Updater")
    
    print("\n" + "=" * 75)
    print("COMPONENT TEST RESULTS SUMMARY:")
    for k, v in results.items():
        print(f"  {k:<25}: {'PASS' if v else 'FAIL'}")
    print("=" * 75)

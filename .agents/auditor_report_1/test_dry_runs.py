import os
import sys
import subprocess
import time

def get_excel_count():
    cmd = ["powershell", "-Command", "(Get-Process EXCEL -ErrorAction SilentlyContinue).Count"]
    res = subprocess.run(cmd, capture_output=True, text=True)
    out = res.stdout.strip()
    return int(out) if out and out.isdigit() else 0

def run_script_benchmark(script_name, args=[]):
    print(f"\n--- Testing Component: {script_name} {' '.join(args)} ---")
    excel_before = get_excel_count()
    t0 = time.time()
    cmd = [sys.executable, f"d:/Alpha/Scripts/{script_name}"] + args
    res = subprocess.run(cmd, cwd="d:/Alpha", capture_output=True, text=True)
    t1 = time.time()
    excel_after = get_excel_count()
    duration = t1 - t0
    
    print(f"  Exit Code: {res.returncode}")
    print(f"  Duration: {duration:.2f}s")
    print(f"  EXCEL.EXE before: {excel_before} | after: {excel_after}")
    if res.returncode != 0:
        print(f"  STDERR: {res.stderr[:300]}")
    else:
        print(f"  STDOUT summary: {res.stdout.strip().splitlines()[-1] if res.stdout.strip() else 'No output'}")
    return res.returncode == 0 and (excel_after == 0)

if __name__ == '__main__':
    print("="*80)
    print("  INDEPENDENT DRY-RUN VERIFICATION BENCHMARK")
    print("="*80)
    
    components = [
        ("sort_dashboard.py", []),
        ("build_archives.py", []),
        ("update_html.py", []),
        ("daily.py", ["--skip-prod", "--skip-wip", "--skip-git"])
    ]
    
    all_ok = True
    for sname, sargs in components:
        ok = run_script_benchmark(sname, sargs)
        if not ok:
            all_ok = False
            
    print("\n" + "="*80)
    if all_ok:
        print("  ALL DRY-RUN BENCHMARKS PASSED SUCCESSFULLY (0 COM LEAKS, 0 ERRORS)")
    else:
        print("  SOME BENCHMARKS FAILED")
    print("="*80)

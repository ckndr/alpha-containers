import subprocess
import time
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

def count_excel_processes():
    try:
        res = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE', '/NH'], capture_output=True, text=True)
        if 'EXCEL.EXE' in res.stdout:
            lines = [l for l in res.stdout.strip().splitlines() if 'EXCEL.EXE' in l]
            return len(lines)
        return 0
    except Exception:
        return 0

print("=" * 80)
print("RUNNING OBJECTIVE 2 DRY RUN BENCHMARKS")
print("=" * 80)

tests = [
    ("sort_dashboard.py", ["python", "Scripts/sort_dashboard.py"]),
    ("build_archives.py", ["python", "Scripts/build_archives.py"]),
    ("update_html.py", ["python", "Scripts/update_html.py"]),
    ("daily.py (--skip-prod --skip-wip --skip-git)", ["python", "Scripts/daily.py", "--skip-prod", "--skip-wip", "--skip-git"])
]

results = []

for name, cmd in tests:
    excel_before = count_excel_processes()
    t0 = time.perf_counter()
    proc = subprocess.run(cmd, cwd="d:\\Alpha", capture_output=True, text=True, encoding='utf-8', errors='replace')
    t1 = time.perf_counter()
    excel_after = count_excel_processes()
    runtime = t1 - t0
    
    status = "SUCCESS" if proc.returncode == 0 else f"FAILED (code {proc.returncode})"
    leak = "CLEAN (0 Leaks)" if excel_after == excel_before else f"LEAK (+{excel_after - excel_before})"
    
    print(f"Component: {name}")
    print(f"  Exit Code:    {proc.returncode} ({status})")
    print(f"  Runtime:      {runtime:.2f}s")
    print(f"  EXCEL Before: {excel_before}, After: {excel_after} -> {leak}")
    if proc.returncode != 0:
        print("  STDERR output:")
        print(proc.stderr[:500])
    print("-" * 80)
    
    results.append({
        "name": name,
        "exit_code": proc.returncode,
        "runtime": runtime,
        "excel_before": excel_before,
        "excel_after": excel_after,
        "leak": leak,
        "stdout_tail": proc.stdout[-300:] if proc.stdout else ""
    })

print("=" * 80)
print("BENCHMARK EXECUTION SUMMARY")
print("=" * 80)
for r in results:
    print(f"{r['name']:45s} | Code: {r['exit_code']} | Time: {r['runtime']:6.2f}s | Leak: {r['leak']}")
print("=" * 80)

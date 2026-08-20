import os
import glob
import re

print("="*80)
print("  PROHIBITED PATTERN & INTEGRITY FORENSIC SCAN")
print("="*80)

scripts = glob.glob("d:/Alpha/Scripts/*.py")
prohibited_findings = []

for s in scripts:
    with open(s, "r", encoding="utf-8", errors="replace") as f:
        content = f.read()
        lines = content.splitlines()

    # Pattern 1: Empty facade functions (pass only or return constant without logic)
    for i, line in enumerate(lines):
        if re.match(r"^\s*def\s+\w+\(.*\):\s*$", line):
            fn_name = line.strip()
            # check next lines
            next_lines = [l.strip() for l in lines[i+1:i+5] if l.strip() and not l.strip().startswith("#") and not l.strip().startswith('"""')]
            if next_lines and next_lines[0] in ["pass", "return True", "return False", "return 0", "return None", "raise NotImplementedError"]:
                if len(next_lines) == 1 or (len(next_lines) > 1 and next_lines[1].startswith("def ")):
                    prohibited_findings.append((s, i+1, f"Potential facade function: {fn_name} -> {next_lines[0]}"))

    # Pattern 2: Hardcoded test mocks or faked bypasses
    for i, line in enumerate(lines):
        if "mock" in line.lower() or "fake" in line.lower() or "bypass" in line.lower() or "dummy" in line.lower():
            # filter out comments
            if not line.strip().startswith("#"):
                prohibited_findings.append((s, i+1, f"Suspicious keyword: {line.strip()}"))

print(f"Scanned {len(scripts)} production scripts.")
print(f"Flagged items: {len(prohibited_findings)}")
for s, lineno, desc in prohibited_findings:
    print(f"  [{os.path.basename(s)}:{lineno}] {desc}")

if not prohibited_findings:
    print("  [CLEAN] No facade implementations, hardcoded test stubs, or faked bypasses found in Scripts/*.py")

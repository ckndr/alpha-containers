with open("Scripts/update_html.py", "r", encoding="utf-8", errors="ignore") as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "timestamp_iso" in line:
        print(f"timestamp_iso at line {i+1}: {line.strip()}")
    if "def inject_block" in line:
        print(f"def inject_block at line {i+1}: {line.strip()}")

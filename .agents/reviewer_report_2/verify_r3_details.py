import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open("Tubex.html", "r", encoding="utf-8", errors="ignore") as f:
    html_lines = f.readlines()

print("Total lines in Tubex.html:", len(html_lines))

# Check escapeHtml definition
for i, line in enumerate(html_lines):
    if "function escapeHtml" in line:
        print(f"\nescapeHtml definition around line {i+1}:")
        for j in range(max(0, i-2), min(len(html_lines), i+10)):
            print(f"  {j+1}: {html_lines[j].rstrip()}")
        break

# Check Orders table escaping
print("\nOrders table escaping lines:")
for i, line in enumerate(html_lines):
    if "escapeHtml(o.customer)" in line or "escapeHtml(o.product)" in line:
        print(f"  {i+1}: {line.strip()}")

# Check FG stock escaping
print("\nFG stock escaping lines:")
for i, line in enumerate(html_lines):
    if "escapeHtml(r.product)" in line or "escapeHtml(r.remarks)" in line:
        print(f"  {i+1}: {line.strip()}")

# Check Customer report filter / data-month
print("\nCustomer report / filter buttons lines:")
for i, line in enumerate(html_lines):
    if "data-month" in line and "escapeHtml" in line:
        print(f"  {i+1}: {line.strip()}")

# Check controllerchange listener
print("\ncontrollerchange listener lines:")
for i, line in enumerate(html_lines):
    if "controllerchange" in line:
        print(f"  {i+1}: {line.strip()}")

# Check sw.js
with open("sw.js", "r", encoding="utf-8", errors="ignore") as f:
    sw_lines = f.readlines()

print("\nTotal lines in sw.js:", len(sw_lines))
for i, line in enumerate(sw_lines):
    if "status === 200" in line or "startsWith('http')" in line or "skipWaiting" in line or "clients.claim" in line or "ASSETS" in line or "return caches.match" in line:
        print(f"  sw.js line {i+1}: {line.strip()}")

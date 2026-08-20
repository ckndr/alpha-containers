import re

with open("d:/Alpha/Tubex.html", "r", encoding="utf-8") as f:
    html = f.read()

with open("d:/Alpha/sw.js", "r", encoding="utf-8") as f:
    sw = f.read()

print("="*80)
print("  WEB DASHBOARD & PWA SECURITY FORENSIC AUDIT")
print("="*80)

# Check 1: escapeHtml function presence and usage
has_escape_fn = "function escapeHtml(str)" in html
print(f"1. Global escapeHtml() helper present: {has_escape_fn}")

# Check 2: innerHTML assignments and escaping
inner_html_assignments = re.findall(r"\.innerHTML\s*=\s*`([^`]+)`", html)
print(f"2. Found {len(inner_html_assignments)} template literal innerHTML assignments.")

unescaped_interpolations = []
for block in inner_html_assignments:
    interpolations = re.findall(r"\$\{([^}]+)\}", block)
    for interp in interpolations:
        interp_clean = interp.strip()
        # If it contains text fields (customer, product, remarks, name) without escapeHtml or number formatting / boolean / ternary
        if any(w in interp_clean for w in ['customer', 'product', 'remarks', 'item.name', 'row.name', 'cat_name']) and 'escapeHtml' not in interp_clean:
            unescaped_interpolations.append(interp_clean)

print(f"3. Unescaped potential user text interpolations in innerHTML: {len(unescaped_interpolations)}")
if unescaped_interpolations:
    for u in unescaped_interpolations:
        print(f"   [WARN] Unescaped: {u}")
else:
    print("   [PASS] All customer, product, remarks, and name fields pass through escapeHtml().")

# Check 4: Service Worker HTTP 200 check
sw_http_200 = "response.status === 200" in sw
print(f"4. Service worker HTTP 200 cache guard: {sw_http_200}")

# Check 5: Service worker scheme check
sw_scheme = "startsWith('http')" in sw
print(f"5. Service worker URL scheme validation: {sw_scheme}")

# Check 6: ISO-8601 timestamp parsing
has_iso_parse = "new Date(DASH_DATA.timestamp_iso)" in html
print(f"6. Standard ISO-8601 timestamp evaluation: {has_iso_parse}")

# Check 7: Offline pre-cache assets
has_index = "./index.html" in sw
has_tubex = "./Tubex.html" in sw
print(f"7. PWA root navigation & fallback pre-cached: index.html={has_index}, Tubex.html={has_tubex}")

import os, sys, datetime, re
import pandas as pd

print("=== 1. Testing escapeHtml snippet ===")
def escape_html(s):
    if s is None:
        return ""
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#39;")
    )

test_str = '<script>alert("XSS & \'attack\'")</script>'
escaped = escape_html(test_str)
print("Escaped string:", escaped)
assert "<" not in escaped and ">" not in escaped and '"' not in escaped and "'" not in escaped
assert "&amp;" in escaped

print("=== 2. Testing parse_date snippet ===")
def parse_date(date_raw):
    if pd.isna(date_raw): return None
    if isinstance(date_raw, (datetime.datetime, datetime.date)): 
        return date_raw if isinstance(date_raw, datetime.date) else date_raw.date()
    try:
        return pd.to_datetime(date_raw, dayfirst=True, errors='coerce').date()
    except Exception:
        return None

d1 = parse_date("06/08/2026") # 6th of August 2026
print("Parsed '06/08/2026':", d1)
assert d1 == datetime.date(2026, 8, 6)

d2 = parse_date(datetime.datetime(2026, 8, 14, 10, 30))
print("Parsed datetime:", d2)
assert d2 == datetime.date(2026, 8, 14)

print("=== 3. Testing coverage ratio check ===")
xls_items = {501, 502, 503}
excel_ids = {i: i+2 for i in range(501, 530)} # 29 items
coverage_ratio = len(xls_items) / max(len(excel_ids), 1)
print(f"Coverage ratio: {coverage_ratio:.1%}")
assert coverage_ratio < 0.70

print("=== 4. Testing fuzzy alias resolution ===")
from difflib import get_close_matches
ALIASES = {
    ("shield baby zinc paste 50g", "25mm"): ("Shield Baby Zinc Paste 50g", 2901),
    ("dettol cream 30g", "19mm"): ("Dettol Antiseptic Cream 30g", 2902),
}

name_raw = "shield baby zinc paste 50gm"
dia_raw = "25mm"
candidate_keys = [k[0] for k in ALIASES.keys() if k[1] == dia_raw]
matches = get_close_matches(name_raw.lower().strip(), candidate_keys, n=1, cutoff=0.85)
print("Candidate keys:", candidate_keys)
print("Matches found:", matches)
if matches:
    resolved = ALIASES.get((matches[0], dia_raw))
    print("Resolved to:", resolved)
    assert resolved[1] == 2901

print("\nALL CODE SNIPPETS PASSED VERIFICATION!")

import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import fitz

doc = fitz.open(r'd:\Alpha\TDS.pdf')

# Print ALL pages - page list first, then specific pages
print("=== TABLE OF CONTENTS ===")
for i in range(len(doc)):
    text = doc[i].get_text()
    # Get first non-empty meaningful line as title
    lines = [l.strip() for l in text.split('\n') if l.strip() and len(l.strip()) > 3]
    # Look for product codes
    title_lines = []
    for l in lines:
        if any(code in l for code in ['422', '400', 'LUBRIMET', 'ALULIQUID', 'SCHEKOSOL', 'SAPILUB', 'SunAltec', 'SECTION']):
            title_lines.append(l)
    if title_lines:
        print(f"  Page {i+1}: {title_lines[0][:100]}")
    else:
        print(f"  Page {i+1}: {lines[0][:100] if lines else '(empty)'}")

# Now print specific pages we need:
# Page 3: LUBRIMET GR8 product data
# Pages for ALULIQUID 13
print("\n\n=== DETAILED: Pages 3-11 (LUBRIMET GR8) ===")
for i in range(2, min(11, len(doc))):
    text = doc[i].get_text()
    text_upper = text.upper()
    if 'LUBRIMET' in text_upper:
        print(f'\n--- Page {i+1} ---')
        print(text[:3000])

print("\n\n=== DETAILED: ALULIQUID 13 pages ===")
for i in range(len(doc)):
    text = doc[i].get_text()
    text_upper = text.upper()
    if 'ALULIQUID' in text_upper:
        print(f'\n--- Page {i+1} ---')
        print(text[:3000])

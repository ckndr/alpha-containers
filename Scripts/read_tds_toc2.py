import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import fitz

doc = fitz.open(r'd:\Alpha\TDS.pdf')

print("=== TABLE OF CONTENTS ===")
for i in range(len(doc)):
    text = doc[i].get_text()
    lines = [l.strip() for l in text.split('\n') if l.strip() and len(l.strip()) > 3]
    # Look for product identifiers
    product_lines = []
    for l in lines[:15]:  # Check first 15 lines
        if any(code in l.upper() for code in ['422', '400', 'LUBRIMET', 'ALULIQUID', 'SCHEKOSOL', 'SAPILUB', 'SUNALTEC', 'SECTION 1', 'SECTION 2']):
            product_lines.append(l)
    if product_lines:
        print(f"  Page {i+1}: {' | '.join(product_lines[:3])}")
    else:
        first_line = lines[0][:120] if lines else '(empty)'
        print(f"  Page {i+1}: {first_line}")

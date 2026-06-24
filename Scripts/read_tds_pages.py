import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import fitz

doc = fitz.open(r'd:\Alpha\TDS.pdf')

# Print pages 1-10 where we expect to find LUBRIMET, ALULIQUID, and INT PROT
for i in range(min(15, len(doc))):
    text = doc[i].get_text()
    # Only print first 15 pages to find the product specs
    print(f'\n{"="*60}')
    print(f'Page {i+1}')
    print(f'{"="*60}')
    # Print just the first 50 lines of each page to get the key info
    lines = text.split('\n')
    for line in lines[:80]:
        print(line)
    if len(lines) > 80:
        print(f'... ({len(lines)-80} more lines)')

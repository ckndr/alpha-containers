import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import fitz

doc = fitz.open(r'd:\Alpha\TDS.pdf')

# Print pages 3, 4 (INT BEIGE 400 4 902, SunAltec ink)
for pg in [3, 4, 5]:
    i = pg - 1
    text = doc[i].get_text()
    print(f'\n{"="*60}')
    print(f'Page {pg}')
    print(f'{"="*60}')
    print(text)
